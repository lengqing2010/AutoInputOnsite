using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.Runtime.InteropServices;
using System.Threading;
using System.Configuration;
using AutoInputOnsite;
using Microsoft.VisualBasic;
namespace AutoInputOnsite
{
    public partial class Menu : Form
    {
        public Menu()
        {
            InitializeComponent();
        }
        public AutoImportCsv FormImportCsv;
        public AutoImportNouhinsyo FormNouki;
        public SHDocVw.InternetExplorerMedium Ie;
        private BackgroundWorker BackgroundWorker;


        // Menu_Load
        private void Menu_Load(System.Object sender, System.EventArgs e)
        {
            QuitIE();
        }

        // Menu_Disposed
        private void Menu_Disposed(object sender, System.EventArgs e)
        {
            QuitIE();
        }


        #region 見出＆明細作成

            //見出＆明細作成 Click
            private void btnMs_Click_1(object sender, EventArgs e)
            {
                this.btnAll.Enabled = false;
                this.btnMs.Enabled = false;
                this.btnNouhin.Enabled = false;
                this.cbInsatu.Enabled = false;
                this.cbWeb.Enabled = false;

                StartAllJunbi();
                DoAutoImportCsv();
                QuitIeAndMsgOk();
            }

            //見出＆明細作成 実行
            private void DoAutoImportCsv()
            {
                FormImportCsv = new AutoImportCsv();
                FormImportCsv.Ie = Ie;
                FormImportCsv.IeVisible = this.cbWeb.Checked;
                if (cbHyouji.Checked)
                {
                    FormImportCsv.Show();
                }
                FormImportCsv.DoAll();
                FormImportCsv.Close();
                pb1.Value = 100;
                FormImportCsv.Dispose();
                FormImportCsv = null;
            }

        #endregion

        #region 納期指定発注
            //納期指定発注 ボタンCLICK
            private void btnNouhin_Click(System.Object sender, System.EventArgs e)
            {
                this.btnAll.Enabled = false;
                this.btnMs.Enabled = false;
                this.btnNouhin.Enabled = false;
                this.cbInsatu.Enabled = false;
                this.cbWeb.Enabled = false;

                StartAllJunbi();
                AutoImportNouhinsyo();
                QuitIeAndMsgOk();
            }
            //納期指定発注 実行
            private void AutoImportNouhinsyo()
            {
                FormNouki = new AutoImportNouhinsyo();
                FormNouki.Ie = Ie;
                try
                {
                    FormNouki.IeVisible = this.cbWeb.Checked;
                    if (cbHyouji.Checked)
                    {
                        FormNouki.Show();
                    }
                }
                catch (Exception)
                {
                }
                FormNouki.insatu = this.cbInsatu.Checked;
                FormNouki.DoAll();
                FormNouki.Close();
                FormNouki.Dispose();
                FormNouki = null;
                pb2.Value = 100;
            }
        #endregion

        #region 見出＆明細作成  AND  納期指定発注 ALL
            /// 見出＆明細作成  AND  納期指定発注 ALL
            private void btnAll_Click_1(System.Object sender, System.EventArgs e)
            {
                this.btnAll.Enabled = false;
                this.btnMs.Enabled = false;
                this.btnNouhin.Enabled = false;
                this.cbInsatu.Enabled = false;
                this.cbWeb.Enabled = false;

                //見出＆明細作成
                StartAllJunbi();
                DoAutoImportCsv();
                //納期指定発注
                StartAllJunbi();
                pb1.Value = 100;
                AutoImportNouhinsyo();
                QuitIeAndMsgOk();
            }
        #endregion

        #region 共通

            /// <summary>
            /// Start 準備 
            ///   1.Close Ie
            ///   2.Init BackgroundWorker
            ///   3.Init Timer 
            /// </summary>
            /// <remarks></remarks>
            private void StartAllJunbi()
            {
                CloseIE();
            re:
                try
                {
                    Ie = null;
                }
                catch (Exception)
                {
                    System.Threading.Thread.Sleep(2000);
                    goto re;
                }
                Ie = new SHDocVw.InternetExplorerMedium();
                if (ReferenceEquals(BackgroundWorker, null))
                {
                    BackgroundWorker = new BackgroundWorker();
                    BackgroundWorker.DoWork += BackgroundWorker_DoWork;
                    BackgroundWorker.RunWorkerAsync();
                }
                else
                {
                    if (!BackgroundWorker.IsBusy)
                    {
                        BackgroundWorker.RunWorkerAsync();
                    }
                    else
                    {
                        BackgroundWorker.Dispose();
                        BackgroundWorker = new BackgroundWorker();
                        BackgroundWorker.DoWork += BackgroundWorker_DoWork;
                        BackgroundWorker.RunWorkerAsync();
                    }
                }

                Timer1.Stop();
                Timer1.Start();
                pb1.Value = 0;
                pb2.Value = 0;
            }

            //IEXPLORE 閉じる
            public void CloseIE()
            {
                try
                {
                    System.Diagnostics.Process[] myProcesses;
                    myProcesses = System.Diagnostics.Process.GetProcessesByName("IEXPLORE");
                    foreach (System.Diagnostics.Process instance in myProcesses)
                    {
                        instance.CloseMainWindow();
                    }
                }
                catch (Exception)
                {
                }
            }

            /// <summary>
            /// 完了
            /// </summary>
            /// <remarks></remarks>
            private void QuitIeAndMsgOk()
            {
                QuitIE();
                this.btnAll.Enabled = true;
                this.btnMs.Enabled = true;
                this.btnNouhin.Enabled = true;
                this.cbInsatu.Enabled = true;
                this.cbWeb.Enabled = true;

                MessageBox.Show("完了");
            }
            
            //Quit Ie
            private void QuitIE()
            {
                try
                {
                    Ie.Quit();
                }
                catch (Exception)
                {
                }
            }

            // Timer1
            private void Timer1_Tick_1(System.Object sender, System.EventArgs e)
            {
                if (FormImportCsv != null)
                {
                    pb1.Value = FormImportCsv.ProBar;
                }
                if (FormNouki != null)
                {
                    pb2.Value = FormNouki.ProBar;
                }
            }

            // Windows Object Something
            private void BackgroundWorker_DoWork(object sender, System.ComponentModel.DoWorkEventArgs e)
            {
                Com com = new Com("");
                com.WindowsInputThread();
            }

        #endregion

    }
}