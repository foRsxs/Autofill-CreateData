using System;
//using System.Collections.Generic;
using System.ComponentModel;
//using System.Data;
using System.Drawing;
//using System.Linq;
using System.Text;
using System.Windows.Forms;
using mshtml;
using System.IO;
using System.Runtime.InteropServices;
using WindowsInput;
using System.Net;
using System.Net.Cache;
using System.Text.RegularExpressions;
using System.Diagnostics;
using System.Collections.Generic;

//Подключение параллельный вычислений
using Excel = Microsoft.Office.Interop.Excel;
using Microsoft.Win32;
using System.Security.Principal;




namespace Autofill
{


    public partial class Form1 : Form
    {
        Properties.Settings ps = Properties.Settings.Default;
        
        //System.Windows.Forms.Timer timerClicker1;
        System.Windows.Forms.Timer timerClicker2;
        //  bool time_to_do = true;
        Options op;
        bool check, err;
        bool work_o = false;
        int kol = 0, kol_error = 0; //j_login = 0, j_err = 0, j_loadmask = 0, j_messagebox = 0;

        delegate void SetTextCallback(string text);
        InputSimulator input = new InputSimulator();

        public string XMLFileName = AppDomain.CurrentDomain.BaseDirectory + "\\base.xml";
        
        
        public string date_form;
        Excel.Application xlApp;
        Excel.Workbook xlWorkbook;
        Excel._Worksheet xlWorksheet;
        Excel.Range xlRange;
        Dictionary<string, WindowsInput.Native.VirtualKeyCode> digitalKeyArray = new Dictionary<string, WindowsInput.Native.VirtualKeyCode>();

        static object locker = new object();

        [DllImport("user32.dll", SetLastError = true)]
        static extern IntPtr FindWindowEx(IntPtr hwndParent, IntPtr hwndChildAfter, string lpszClass, string lpszWindow);

        [DllImport("user32.dll", EntryPoint = "FindWindow", SetLastError = true)]
        private static extern IntPtr FindWindow(string lpClassName, string lpWindowName);

        [DllImport("user32.dll", CharSet = CharSet.Auto)]
        static extern IntPtr SendMessage(IntPtr hWnd, UInt32 Msg, IntPtr wParam, IntPtr lParam);

        [DllImport("user32.dll")]
        static extern void mouse_event(int dwFlags, int dx, int dy, int dwData, int dwExtraInfo);

        [Flags]
        public enum MouseEventFlags
        {
            LEFTDOWN = 0x00000002,
            LEFTUP = 0x00000004,
            MIDDLEDOWN = 0x00000020,
            MIDDLEUP = 0x00000040,
            MOVE = 0x00000001,
            ABSOLUTE = 0x00008000,
            RIGHTDOWN = 0x00000008,
            RIGHTUP = 0x00000010
        }




        public Form1()
        {
            InitializeComponent();
        }


        private void Form1_Load(object sender, EventArgs e)
        {
            this.WindowState = FormWindowState.Maximized;
            splitContainer2.SplitterDistance = (this.Width / 100) * 80;


            splitContainer1.Panel2Collapsed = true;
            splitContainer1.Panel2.Hide();

            textBox1.ScrollBars = ScrollBars.Vertical;
 
            webBrowser1.ScriptErrorsSuppressed = true;
            webBrowser1.ProgressChanged += new WebBrowserProgressChangedEventHandler(webState);


            webBrowser1.Navigated +=
            new WebBrowserNavigatedEventHandler(
                (object _sender, WebBrowserNavigatedEventArgs args) => {
                    Action<HtmlDocument> blockAlerts = (HtmlDocument d) => {
                        HtmlElement h = d.GetElementsByTagName("head")[0];
                        HtmlElement s = d.CreateElement("script");
                        IHTMLScriptElement _e = (IHTMLScriptElement)s.DomElement;
                        _e.text = "window.alert=function(){};";
                        h.AppendChild(s);
                    };
                    WebBrowser b = _sender as WebBrowser;
                    blockAlerts(b.Document);
                    for (int i = 0; i < b.Document.Window.Frames.Count; i++)
                        try { blockAlerts(b.Document.Window.Frames[i].Document); }
                        catch (Exception) { };
                }
            );

            webBrowser1.Navigate(ps.url);

            /*
            timerClicker2 = new System.Windows.Forms.Timer();
            timerClicker2.Interval = 30000;
            timerClicker2.Tick += new EventHandler(timerClicker_Tick2);
            timerClicker2.Start();*/

            textBox1.Text = DateTime.Now.ToShortTimeString() + " - Программа по автоматическому заполнению запущена.";

            if (IsUserAdministrator()) SetText("Запущен в режиме Администратора");
            else SetText("Запущен в обычном режиме");

            var appName = Process.GetCurrentProcess().ProcessName + ".exe";
            SetIE8KeyforWebBrowserControl(appName);
            

            if (ps.status) statusStrip2.Visible = true;
            else statusStrip2.Visible = false;




            //Заполняем массив для кнопок Цифр (при заполнении дат, воспринимаются только эти кнопки, эмуляция не помогает)
   
            digitalKeyArray.Add("0", WindowsInput.Native.VirtualKeyCode.VK_0);
            digitalKeyArray.Add("1", WindowsInput.Native.VirtualKeyCode.VK_1);
            digitalKeyArray.Add("2", WindowsInput.Native.VirtualKeyCode.VK_2);
            digitalKeyArray.Add("3", WindowsInput.Native.VirtualKeyCode.VK_3);
            digitalKeyArray.Add("4", WindowsInput.Native.VirtualKeyCode.VK_4);
            digitalKeyArray.Add("5", WindowsInput.Native.VirtualKeyCode.VK_5);
            digitalKeyArray.Add("6", WindowsInput.Native.VirtualKeyCode.VK_6);
            digitalKeyArray.Add("7", WindowsInput.Native.VirtualKeyCode.VK_7);
            digitalKeyArray.Add("8", WindowsInput.Native.VirtualKeyCode.VK_8);
            digitalKeyArray.Add("9", WindowsInput.Native.VirtualKeyCode.VK_9);

            настройкиToolStripMenuItem_Click(sender,e);

        }



    private void webState(object sender, WebBrowserProgressChangedEventArgs e)
        {
            long percent = (e.CurrentProgress * 100)/ e.MaximumProgress;
            toolStripProgressBar2.Value=Convert.ToInt32(percent);
        }


        private void cliclAlertDialog()
        {
            IntPtr hwnd = FindWindow("#32770", "Message from webpage");
            hwnd = FindWindowEx(hwnd, IntPtr.Zero, "Button", "OK");
            uint message = 0xf5;
            SendMessage(hwnd, message, IntPtr.Zero, IntPtr.Zero);
        }



        public bool IsUserAdministrator()
        {
            bool isAdmin;
            try
            {
                WindowsIdentity user = WindowsIdentity.GetCurrent();
                WindowsPrincipal principal = new WindowsPrincipal(user);
                isAdmin = principal.IsInRole(WindowsBuiltInRole.Administrator);
            }
            catch (UnauthorizedAccessException ex)
            {
                isAdmin = false;
            }
            catch (Exception ex)
            {
                isAdmin = false;
            }
            return isAdmin;
        }


        private void SetIE8KeyforWebBrowserControl(string appName)
        {
            RegistryKey Regkey = null;
            try
            {

                //For 64 bit Machine 
                if (Environment.Is64BitOperatingSystem)
                    Regkey = Microsoft.Win32.Registry.LocalMachine.OpenSubKey(@"SOFTWARE\\Wow6432Node\\Microsoft\\Internet Explorer\\MAIN\\FeatureControl\\FEATURE_BROWSER_EMULATION", true);
                else  //For 32 bit Machine 
                    Regkey = Microsoft.Win32.Registry.LocalMachine.OpenSubKey(@"SOFTWARE\\Microsoft\\Internet Explorer\\Main\\FeatureControl\\FEATURE_BROWSER_EMULATION", true);

              
                //If the path is not correct or 
                //If user't have priviledges to access registry 
                if (Regkey == null)
                {
                    SetText("Ошибка установки настроек - Не найден адрес");
                    return;
                }

                string FindAppkey = Convert.ToString(Regkey.GetValue(appName));
                // SetText("Текущий ключ браузера: " + FindAppkey);

                //Check if key is already present 
                if (FindAppkey == ps.bro_version)
                {
                    SetText("Требуемые настройки установлены ранее");
                    Regkey.Close();
                    return;
                }
                else Regkey.SetValue(appName, unchecked(ps.bro_version), RegistryValueKind.DWord);

                

                //check for the key after adding 
                FindAppkey = Convert.ToString(Regkey.GetValue(appName));

                if (FindAppkey == ps.bro_version)
                    MessageBox.Show("Настройки успешно установлены");
                else
                    MessageBox.Show("Ошибка установки найстроек, Ref: " + FindAppkey);


            }
            catch (Exception ex)
            {
                MessageBox.Show("Ошибка установки настроек");
                MessageBox.Show(ex.Message);
            }
            finally
            {
                //Close the Registry 
                if (Regkey != null)
                    Regkey.Close();
            }


        }

        public static void LeftClick(int x, int y)
        {
            Cursor.Position = new System.Drawing.Point(x, y);
            mouse_event((int)(MouseEventFlags.LEFTDOWN), 0, 0, 0, 0);
            mouse_event((int)(MouseEventFlags.LEFTUP), 0, 0, 0, 0);
        }

        void timerClicker_Tick2(object sender, EventArgs e)
        { 
            check_login();
            check_err();
        }


        /// <summary>
        /// подождать указанное время
        /// </summary>
        /// <param name="seconds"> время в секундах </param>
        private void Wait(double seconds)
        {
            int ticks = System.Environment.TickCount + (int)Math.Round(seconds * 1000.0);
            
            while (System.Environment.TickCount < ticks)
            {
                Application.DoEvents();
            }
        }

        //Ожидание готовности браузера
        private void wbWait()
        {
            string ajax_load = "none";

            do
            {
                Wait(ps.w2);
               
                if (webBrowser1.Document.GetElementById("AjaxLoad") != null)
                ajax_load = webBrowser1.Document.GetElementById("AjaxLoad").GetAttribute("style").ToString();
            } //webBrowser1.IsBusy || webBrowser1.ReadyState != WebBrowserReadyState.Complete ||
            while (webBrowser1.IsBusy || webBrowser1.ReadyState != WebBrowserReadyState.Complete || ajax_load.Contains("display: block;"));
        }

        private void обновитьToolStripMenuItem_Click(object sender, EventArgs e)
        {
           // webControl1.Source = new Uri(ps.url);
            webBrowser1.Navigate(ps.url);
        }

        private void Excel_close()
        {
            // Cleanup
            GC.Collect();
            GC.WaitForPendingFinalizers();

            if (xlRange != null)
            {
                Marshal.FinalReleaseComObject(xlRange);
                xlRange = null;
            }

            if (xlWorksheet != null)
            {
                Marshal.FinalReleaseComObject(xlWorksheet);
                xlWorksheet = null;
            }
            if (xlWorkbook != null)
            {
                xlApp.DisplayAlerts = false;
                xlWorkbook.Close(true, Type.Missing, Type.Missing);
                Marshal.FinalReleaseComObject(xlWorkbook);
                xlWorkbook = null;
            }
            if (xlApp != null)
            {
                xlApp.Quit();
                Marshal.FinalReleaseComObject(xlApp);
                xlApp = null;
            }


            Process[] processes = Process.GetProcessesByName("EXCEL");
            foreach (Process p in processes)
            {
                p.Kill();
            }


        }

        // Функция выоплнения Javascript
        private void JSinvoke(string script, string name_script)
        {
            HtmlDocument doc = webBrowser1.Document;
            HtmlElement head = doc.GetElementsByTagName("head")[0];
            HtmlElement scriptEl = doc.CreateElement("script");

            scriptEl.SetAttribute("text", script);
           // IHTMLScriptElement element = (IHTMLScriptElement)scriptEl.DomElement;
           //element.text = script; // example "function sayHello() { alert('hello') }"
            head.AppendChild(scriptEl);
            doc.InvokeScript(name_script);
        }


        //Возращает событие о готовности браузера
        private bool wb_readystate()
        {
            if (webBrowser1.ReadyState == WebBrowserReadyState.Complete)
                return true;
            else return false;
        }

        //Устанавливаем статус
        private void statusbar(string status)
        {
            toolStripStatusLabel1.Text = status;
        }


        //Эмуляция Клика через javascript
        public void JsFireEvent(string getElementQuery, string eventName)
        {
            if (wb_readystate() && !string.IsNullOrEmpty(getElementQuery))
            {
              /*  webControl1.ExecuteJavascript(@"
                            function fireEvent(element,event) {
                                var evt = document.createEvent('HTMLEvents');
                                evt.initEvent(event, true, true ); // event type,bubbling,cancelable
                                element.dispatchEvent(evt);                                 
                            }
                            " + String.Format("fireEvent({0}, '{1}');", getElementQuery, eventName));*/

                JSinvoke(@"
                            function fireEvent(element,event) {
                                var evt = document.createEvent('HTMLEvents');
                                evt.initEvent(event, true, true ); // event type,bubbling,cancelable
                                element.dispatchEvent(evt);                                 
                            }
                            " + String.Format("fireEvent({0}, '{1}');", getElementQuery, eventName), eventName);

            }
        }

        // Click by ID
        private void click_by_id(string id)
        {
            HtmlElement fbLink = webBrowser1.Document.GetElementById(id);
            if (fbLink != null)
                fbLink.InvokeMember("click");
            else SetText("Не найдена кнопка для клика с ID: " + id);
        }

        //Focus by ID
        private void focusById(string id)
        {
            HtmlElement _el = webBrowser1.Document.GetElementById(id);
            if (_el != null)
                _el.Focus();
            else SetText("focusById - Не найден элемент");
        }

        //Find By Inner html
        private void clickByInnerHtml(string _tag, string _inner)
        {
            HtmlElementCollection col = webBrowser1.Document.GetElementsByTagName(_tag);

            foreach (HtmlElement item in col)
            {
                if (item.InnerText == _inner)
                {
                    item.InvokeMember("Click");
                }
            }
        }

        //Поиск элемента по аттрибуту
        private HtmlElement findByAttribute(string _tag, string attr_name, string attr_value)
        {
            HtmlElementCollection col = webBrowser1.Document.GetElementsByTagName(_tag);
            HtmlElement wanted = null;           

            foreach (HtmlElement item in col)
            {
                if (item.GetAttribute(attr_name).Contains(attr_value))
                {
                    wanted = item;
                    break;
                }
            }

            return wanted;
        }

        //Эмуляации нажатий цифровых клавиш
        private void simulateDigitKey(string digit)
        {
            foreach (char dig in digit)
            {
                if (char.IsDigit(dig)) input.Keyboard.KeyPress(digitalKeyArray[dig.ToString()]);
                Wait(0.1);
            }
        }

        //Вставить значение по ID элемента
        private void insert_value_byid(string id, string value)
        {
            JSinvoke("function insert_value_byid() { $('[id *= " + id + "]').val( '" + value + "' ); }", "insert_value_byid");
        }


        //Вставить значение по NAME элемента
        private void insert_value_byName(string name, string value)
        {
            JSinvoke("function insert_value_byName() { $('[name *= " + name + "]').val( '" + value + "' ); }", "insert_value_byName");
            // webControl1.ExecuteJavascript("$(document).ready(function () { $('[name *= " + id + "]').val( '" + value + "' ); });");
        }

        //Если найдена форма входа
        private void check_login()
        {
           

            if (wb_readystate())
            {

                string el = webBrowser1.Document.Title.ToString();

                    if (el.Contains("Выполнить вход") || el.Contains("Главная страница"))
                    {
                        if (el.Contains("Выполнить вход"))
                        {
                            //LeftClick(Convert.ToInt32(ps.login.Split(':')[0].ToString()), Convert.ToInt32(ps.login.Split(':')[1].ToString()));
                            focusById("UserName");
                            Wait(0.3);
                            input.Keyboard.TextEntry(ps.login_text);
                                Wait(0.5);

                            //LeftClick(Convert.ToInt32(ps.pass.Split(':')[0].ToString()), Convert.ToInt32(ps.pass.Split(':')[1].ToString()));
                            focusById("Password");
                            Wait(0.3);
                            input.Keyboard.TextEntry(ps.pass_text);

                            input.Keyboard.KeyPress(WindowsInput.Native.VirtualKeyCode.RETURN);

                            SetText("Произведен Вход на страницу: ");
 
                        }

                        wbWait(); //Ждем когда браузер будет готов

                        webBrowser1.Navigate("https://pol.eisz.kz/app/Human");

                        wbWait();


                    /*
                        if (ps.login_text == "821012401599")
                        {

                            HtmlElementCollection col = webBrowser1.Document.GetElementsByTagName("a");

                            foreach (HtmlElement item in col)
                            {
                                if (item.InnerText.Contains("Шымкентская городская поликлиника №1"))
                                {
                                    item.InvokeMember("Click");
                                     wbWait();
                                }
                            }
                                            

                        }
                        */

                    statusbar("Вход выполнен");
                    }
               

            }
        }


        //Если найдена Ошибка
        private bool check_err()
        {
            string el = "";

            if (wb_readystate())
            {
                dynamic element = webBrowser1.Document.GetElementById("eerWin");
                if (element == null) return false;

                el = element.GetAttribute("classname").ToString();

                if (!el.Contains("x-hide-offsets"))
                {
                    return true;
                } else return false;

            }  else return false;
        }


        //Проверка загрузки TRUE если окно не найдено
        private bool check_messagebox()
        {
            if (wb_readystate())
            {
                try
                {
                    toolStripStatusLabel1.Text = "Поиск открытого окна с кнопкой ОК";

                    HtmlElement h_elem = webBrowser1.Document.GetElementById("messagebox-1001");
                    string el = "";

                    if (h_elem == null) return true; //возравщаем что все ОК

                    if (!h_elem.GetAttribute("class").ToString().Contains("x-hide-offsets")) //если окно показывается
                    {
                        return true;
                    }
                    else return false;
                }
                catch
                {
                    SetText("Warning - Messagebox.");
                }
            }
            return false;
        }


        private bool check_element(string id_elem, string attribute)
        {
            if (wb_readystate())
            {
                try
                {
                   
                    dynamic Win = webBrowser1.Document.GetElementById(id_elem);
                    string el = "";

                    if (Win == null) return false;
                    el = Win.GetAttribute("classname").ToString();

                    if (el.Contains(attribute)) return true;
                    else return false;
                }
                catch
                {
                    SetText("Warning - Check_element.");
                    return false;
                    
                }
            }
            return false;
        }


        private void fill_open()
        {

            Double w1 = ps.w1, w2 = ps.w2, w3 = ps.w3;

            if (work_o && wb_readystate())
            {
                this.Activate();

                OpenFileDialog of = new OpenFileDialog();
                of.Filter = "Excel Files|*.xls;*.xlsx;*.xlsm";

                if (ps.of_dir.Length < 2)  of.InitialDirectory = @"C:\"; //Если путь файла не установлен
                else of.InitialDirectory = ps.of_dir;

                if (of.ShowDialog() == DialogResult.OK)
                {
                    if (of.FileName.Length < 1) return;

                    //Сохраняем путь файла
                    Properties.Settings.Default["of_dir"] = Path.GetDirectoryName(of.FileName);
                    Properties.Settings.Default.Save();

                    xlApp = new Excel.Application();
                    xlWorkbook = xlApp.Workbooks.Open(of.FileName, 0, false, 5, "", "", true, Excel.XlPlatform.xlWindows, "\t", true, false, 0, true, 1, 0);
                    xlWorksheet = (Excel._Worksheet)xlWorkbook.Sheets[1];

                    toolStripStatusLabel1.Text = "Начинаем работу с файлом:"+of.FileName;

                    //These two lines do the magic.
                    xlWorksheet.Columns.ClearFormats();
                    xlWorksheet.Rows.ClearFormats();
                    xlRange = xlWorksheet.UsedRange;


                    int rowCount = xlRange.Rows.Count;
                    int colCount = xlRange.Columns.Count;

                    kol = 0;
                    kol_error = 0;
                    toolStripStatusLabel2.Text = "Выполнено: " + kol.ToString() + " | Ошибки: " + kol_error.ToString();

                   

                    for (int i = 2; i <= rowCount; i++)
                    {

                        string error = "", iin = "";
                        int number_string = 0;

                        // check_login();
                        check_err();

                        if (!work_o) break;



                        webBrowser1.Navigate("https://pol.eisz.kz/app/Human");//Update page
                        wbWait(); //wait 2-3 seconds for loaded page

                        //Журнал исходящих направлений из ПМСП
                        /*  toolStripStatusLabel1.Text = "Проверяем страницу - Журнал ПМСП";
                          string el4 = webBrowser1.Document.Title.ToString();

                          if (!el4.Contains("Журнал исходящих направлений из ПМСП"))
                          {
                              SetText("Открыта другая страница: " + el4);
                              обновитьToolStripMenuItem_Click(this, new EventArgs());

                              while (!wb_readystate())
                              {
                                  Wait(0.3);
                              }

                              while (!webBrowser1.Document.Title.ToString().Contains("Журнал исходящих направлений из ПМСП"))
                              {
                                  check_login();
                              }

                              Wait(3.0);
                          }
                          */


                        //EXCEL Проверяем все ячейки в в текущей строке
                        for (int j = 1; j <= 8; j++)
                        {
                            toolStripStatusLabel1.Text = "Проверка всех ячеек";

                            if (((Excel.Range)xlRange.Cells[i, j]).Value2 == null)
                            { 
                                error = "Строка №" + i + ", содержит пустую(ые) ячейку.";
                                number_string = i;
                            }

                            if (j == 1) //Если первый столбец ИИН
                            {
                                //EXCEL Проверяем ячейку ИИН
                                toolStripStatusLabel1.Text = "Проверка ИИН пациента";
                                iin = ((Excel.Range)xlRange.Cells[i, 1]).Value2.ToString();
                                if (iin.Length != 12)
                                {
                                    error = "Строка №" + i + ", ИИН не 12 знаков.";
                                    number_string = i;
                                }
                            }
         
                        }

                        


                    //Если все ячейки текущей строки без ошибок
                        if (error.Length<1)
                        {

                            string
                                date1 = ((Excel.Range)xlRange.Cells[i, 2]).Value2.ToString(),
                                vrach = ((Excel.Range)xlRange.Cells[i, 3]).Value2.ToString().Trim(),
                                dateNaprav = ((Excel.Range)xlRange.Cells[i, 4]).Value2.ToString(),
                                code_uslugi = ((Excel.Range)xlRange.Cells[i, 5]).Value2.ToString(),
                                zaveduw = ((Excel.Range)xlRange.Cells[i, 6]).Value2.ToString().Trim(),
                                dateAccept = ((Excel.Range)xlRange.Cells[i, 7]).Value2.ToString(),
                                vrachNapr = ((Excel.Range)xlRange.Cells[i, 8]).Value2.ToString().Trim();

                           // SetText("date1:" + date1 + ", dateNaprav:" + dateNaprav);

                            int  medorg = Convert.ToInt32(ps.medorg.ToString()),
                                 mestApp = Convert.ToInt32(ps.medotd.ToString());

                            //Пока не будет найден элемент ввода ИИН ждем
                            while (webBrowser1.Document.GetElementById("criteria") == null)
                            {
                                wbWait();
                                statusbar("Ожидания поля ввода ИИН");
                            }

                            statusbar("Вводим ИИН");

                            focusById("criteria");
                            //webBrowser1.Document.GetElementById("Find").SetAttribute("value", "");

                            Wait(1.0);

                            //Insert IIN
                            input.Keyboard.TextEntry(iin);
                            Wait(1.0);
                            input.Keyboard.KeyPress(WindowsInput.Native.VirtualKeyCode.RETURN);

                            wbWait();

                            //Открываем вкладку Стат карты
                            HtmlElement stat_card = findByAttribute("a", "href", "#t2");
                            if (stat_card != null)
                            {
                                stat_card.InvokeMember("Click");
                                SetText("Переход на вкладку Стат карты");
                            }
                            wbWait();

                            //Создаем элемент для чек открытия окна
                            HtmlElement KartIs = webBrowser1.Document.CreateElement("input");
                            KartIs.Style = "display: none;";
                            KartIs.SetAttribute("value", "false");
                            KartIs.SetAttribute("id", "KartIsOpened");
                            stat_card.AppendChild(KartIs);

                            //Открываем окно направления
                            HtmlElement add_card = findByAttribute("button", "data-bind", "click:addCard5Y");
                            if (add_card != null)
                            {
                                add_card.InvokeMember("Click");
                                SetText("Производим нажатие на кнопку открытия карты");
                            }
                            wbWait();

                            //SetText("Создаем новую стат карту");

                           
                            string statPopup = webBrowser1.Document.GetElementById("KartIsOpened").GetAttribute("value").ToString();
                            

                            Wait(1);
                            wbWait();

                            if (statPopup != "true")
                            {
                                int count = 1;

                               
                                do
                                {
                                    statusbar("Ожидаем открытие стат карты (" + count.ToString() + " сек.) ");
                                    Wait(count);
                                    statPopup = webBrowser1.Document.GetElementById("KartIsOpened").GetAttribute("value").ToString();
                                    count++;

                                    if (count > 5)
                                    {
                                        SetText("Не найдено окно для заполнения (5 циклов)");
                                        count = 0;
                                        break;
                                    }

                                }
                                while (statPopup != "true");

                            }


                            if (statPopup == "true")
                            {
                                statusbar("Стат карта открылась");

                                //TAB -- Адрес
                                //Выбираем город
                                focusById("kart_iscity");

                                input.Keyboard.KeyPress(WindowsInput.Native.VirtualKeyCode.DOWN);

                                if (ps.is_city == false) //Если Село
                                {
                                    Wait(w1);
                                    input.Keyboard.KeyPress(WindowsInput.Native.VirtualKeyCode.DOWN);
                                }

                                input.Keyboard.KeyPress(WindowsInput.Native.VirtualKeyCode.RETURN);
                                Wait(w1);

                                focusById("tabs");

                                Wait(w1);



                                //TAB -- Посещение
                                clickByInnerHtml("a", "Посещения");
                                focusById("kart_cause_kod");
                                Wait(w1);
                                input.Keyboard.KeyPress(WindowsInput.Native.VirtualKeyCode.DOWN);
                                Wait(w1);
                                focusById("poseshs");
                                webBrowser1.Document.InvokeScript("AddPolPosesh"); //добавляем поле

                                string id_date = webBrowser1.Document.GetElementsByTagName("input").GetElementsByName("posesh[0].dt_pos")[0].GetAttribute("id").ToString();


                                focusById(id_date);
                                Wait(w1);
                                //Вставляем дату
                                simulateDigitKey(date1);
                                webBrowser1.Document.InvokeScript("UpdateIshodDoctor");
                                wbWait();
                                /*input.Keyboard.KeyPress(WindowsInput.Native.VirtualKeyCode.TAB);
                                Wait(w1);
                                focusById("posesh[0].vra_uidText");
                                input.Keyboard.TextEntry(vrach);  //ФИО РАЧА
                                Wait(w1);
                                input.Keyboard.KeyPress(WindowsInput.Native.VirtualKeyCode.DOWN);
                                Wait(w1);
                                input.Keyboard.KeyPress(WindowsInput.Native.VirtualKeyCode.RETURN);
                                wbWait();
                                */


                                //TAB - Открываем вкладку Диагноз
                                clickByInnerHtml("a", "Диагноз");
                                Wait(w1);
                                webBrowser1.Document.InvokeScript("AddDiagnoz");
                                Wait(w1);
                                focusById("pol_karty_diag[0].spmkb_id");
                                input.Keyboard.KeyPress(WindowsInput.Native.VirtualKeyCode.VK_Z);
                                input.Keyboard.KeyPress(WindowsInput.Native.VirtualKeyCode.VK_0);
                                input.Keyboard.KeyPress(WindowsInput.Native.VirtualKeyCode.VK_0);
                                input.Keyboard.KeyPress(WindowsInput.Native.VirtualKeyCode.VK_0);
                                wbWait();
                                /*
                                focusById("pol_karty_diag[0].vra_uidText");

                                Wait(w1);
                                input.Keyboard.TextEntry(vrach);
                                Wait(w1);
                                input.Keyboard.KeyPress(WindowsInput.Native.VirtualKeyCode.DOWN);
                                Wait(w1);
                                input.Keyboard.KeyPress(WindowsInput.Native.VirtualKeyCode.RETURN);
                                wbWait();*/


                                //TAB -- Исход
                                clickByInnerHtml("a", "Исход");
                                webBrowser1.Document.GetElementById("kart_ishod_screen_kod").SetAttribute("value", "13");
                                webBrowser1.Document.GetElementById("slobr").SetAttribute("value", "1");
                                Wait(w1);
                                click_by_id("SPOUserEnd");
                                wbWait();

                                //TAB -- Открываем вкладку Направление в КДЦ
                                clickByInnerHtml("a", "Направление в КДЦ");
                                Wait(w1);

                         //Открываем окно направления
                                HtmlElement add_napr = findByAttribute("img", "data-bind", "click: $root.CreateNapr");
                                if (add_napr != null)
                                    add_napr.InvokeMember("Click");
                                wbWait();

                                //Доп. окно направления


                                //Выбираем диагноз направления
                                focusById("Diagnozundefined");
                                HtmlElement z00 = findByAttribute("option", "value", "Z00.0");
                                if (z00 != null)
                                {
                                    z00.InvokeMember("Click");
                                    input.Keyboard.KeyPress(WindowsInput.Native.VirtualKeyCode.DOWN);
                                }
                                wbWait();

                                //Вставляем дату
                                HtmlElement date_napr = findByAttribute("input", "iid", "dt_sent");
                                if (date_napr != null)
                                {
                                    date_napr.Focus();
                                    simulateDigitKey(dateNaprav);
                                }
                                wbWait();

                                //Выбираем Местный АПП
                                input.Keyboard.KeyPress(WindowsInput.Native.VirtualKeyCode.TAB);
                                Wait(w1);
                                for (int k = 0; k < mestApp; k++)
                                {
                                    input.Keyboard.KeyPress(WindowsInput.Native.VirtualKeyCode.DOWN);
                                    wbWait();
                                }
                                

                                //Выбираем Организация в которую направляется МЕТОД1
                      
                                mshtml.IHTMLDocument2 htmlDocument = (mshtml.IHTMLDocument2)webBrowser1.Document.DomDocument;
                                var dropdown = ((IHTMLElement)htmlDocument.all.item("NotAllOkpolu0"));
                                var dropdownItems = (IHTMLElementCollection)dropdown.children;
                                int zz = 0;

                                foreach (IHTMLElement option in dropdownItems)
                                {
                                    var value = option.getAttribute("value").ToString();
                                    if (value.Equals("01ZE"))
                                        break;
                                    zz++;
                                }
                         

                                //Выбираем Организация в которую направляется
                                Wait(w1);
                                HtmlElement org = webBrowser1.Document.GetElementById("NotAllOkpolu0");
                                if (org != null)
                                {
                                    org.Focus();

                                    for (int k = 0; k < zz; k++)
                                    {
                                        input.Keyboard.KeyPress(WindowsInput.Native.VirtualKeyCode.DOWN);
                                        Wait(0.5);
                                    }
                                    input.Keyboard.KeyPress(WindowsInput.Native.VirtualKeyCode.RETURN);
                                    Wait(w1);
                                }
                       


                                //Заполняем таблицу направления
                     

                                string[] arr_code_uslugi = code_uslugi.Split(','); // получаем коды услуг

                                for (int m = 0; m < arr_code_uslugi.Length; m++)
                                {

                                    click_by_id("addProc0");
                                    Wait(w1);

                                    webBrowser1.Document.GetElementById("0type"+m).SetAttribute("value", "true");
                                    focusById("0tarusl"+m);
                                    input.Keyboard.TextEntry(arr_code_uslugi[m].Trim());
                                    Wait(w1);
                                    input.Keyboard.KeyPress(WindowsInput.Native.VirtualKeyCode.RETURN);

                                    Wait(w1);

                                    focusById("0vra_uid_zav"+m+"Text");
                                    input.Keyboard.TextEntry(zaveduw);  //Заведущая
                                    Wait(w1);
                                    input.Keyboard.KeyPress(WindowsInput.Native.VirtualKeyCode.TAB);
                                    Wait(w1);
                                    simulateDigitKey(dateAccept);
                                    Wait(w2);

                                }

                                focusById("vra_uid0Text");
                                input.Keyboard.TextEntry(vrachNapr);  //ФИО РАЧА Направивший
                                Wait(w1);

                                //Ставим галочку
                                HtmlElement check_napr = findByAttribute("input", "data-bind", "checked: Confirm");
                                if (check_napr != null)
                                    check_napr.InvokeMember("Click");
                                wbWait();

                                //Кнопка вернуться
                                HtmlElement return_napr = findByAttribute("input", "data-bind", "click: $root.Close.bind($data, $index(), $data)");
                                if (return_napr != null)
                                    return_napr.InvokeMember("Click");
                                //Сохранить стат карту
                                HtmlElement save_but = findByAttribute("input", "value", "Сохранить и закрыть");
                                if (save_but != null)
                                {
                                    save_but.InvokeMember("Click");
                                }
                                //click_by_id("Cancel");

                                wbWait();
                                Wait(w3);

                                
                                HtmlElement kartError = webBrowser1.Document.GetElementById("KartErrorList");
                                if (kartError.InnerText != null && kartError.InnerText.Trim().Length > 1)
                                {
                                    //kartError.InnerText.Contains("Посещения");
                                    SetText("Портал вернул ошибку: "+ kartError.InnerText);
                                    error = "Eiszk error";
                                }

                                wbWait();
                            }
                            else
                            {
                                SetText("Окно стат. карты не найдено!");
                                error = "Окно стат. карты не найдено!";
                            }




                        }
                        


                        //Закрашиваем
                        if (error.Length < 1)
                        {
                            kol++;
                            SetText("Обработана строка № " + i + " (ИИН: " + iin + ")");
                            ((Excel.Range)xlRange.Cells[i, 1]).Interior.Color = Color.Green;
                        }
                        else
                        {
                            kol_error++;
                            SetText("Строка №: " + i + " не обработалась! Ошибка: "+error);
                          
                            ((Excel.Range)xlRange.Cells[i, 1]).Interior.Color = Color.Red;
                        }

                        xlApp.DisplayAlerts = false;
                        xlWorkbook.Save();

                        toolStripStatusLabel2.Text = "Выполнено: " + kol.ToString() + " | Ошибки: " + kol_error.ToString();

                    }


                    SetText("Подготовка к завершению");
                    toolStripStatusLabel1.Text = "Завершение";

                    xlApp.DisplayAlerts = false;
                    xlWorkbook.Save();

                    Excel_close();

                    SetText("Обработка файла завершена!");
                    MessageBox.Show("Обработка файла завершена!");


                }
               

           

            }
        }



        //Заполняем Открытые
        private void заполToolStripMenuItem_Click(object sender, EventArgs e)
        {
            work_o = true;

            var DD = new Date();
            DD.ShowDialog();

            fill_open();
        }

        private void настройкиToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (op == null || op.IsDisposed)
            {
                op = new Options();
                op.Owner = this;
                op.Show();
            }
            else op.Activate();
        }

        private void progress_set(int value)
        {
            if (!this.InvokeRequired)
            {
                //toolStripProgressBar1.Value = value;
            }
        }

        private void выходToolStripMenuItem_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void сТОПToolStripMenuItem_Click(object sender, EventArgs e)
        {
            _stop();
        }

        private void _stop()
        {
            work_o = false;
          //  work_p = false;
           
            SetText(DateTime.Now.ToShortTimeString() + " - Остановлено.");
            toolStripStatusLabel1.Text = "Остановлено";
        }

        private void оПрограммеToolStripMenuItem_Click(object sender, EventArgs e)
        {
            AboutBox1 ab = new AboutBox1();
            ab.ShowDialog();
        }

        private void check_validation()
        {
            if (CheckForInternetConnection())
            {
                SetText("Проверка Validation Period.");

                DateTime date_valid = new DateTime();
                date_valid = DateTime.ParseExact("2014-08-01 10:10:30.120", "yyyy-MM-dd hh:mm:ss.fff", null);

                
                if (GetNistTime() > date_valid)
                {
                    ps.validation = false;
                }
            }

        }

        public static DateTime GetNistTime()
        {
            DateTime dateTime = DateTime.MinValue;

            try
            {
                HttpWebRequest request = (HttpWebRequest)WebRequest.Create("http://nist.time.gov/actualtime.cgi?lzbc=siqm9b");
                request.Method = "GET";
                request.Accept = "text/html, application/xhtml+xml, */*";
                request.UserAgent = "Mozilla/5.0 (compatible; MSIE 10.0; Windows NT 6.1; Trident/6.0)";
                request.ContentType = "application/x-www-form-urlencoded";
                request.CachePolicy = new RequestCachePolicy(RequestCacheLevel.NoCacheNoStore); //No caching
                HttpWebResponse response = (HttpWebResponse)request.GetResponse();
                if (response.StatusCode == HttpStatusCode.OK)
                {
                    StreamReader stream = new StreamReader(response.GetResponseStream());
                    string html = stream.ReadToEnd();//<timestamp time=\"1395772696469995\" delay=\"1395772696469995\"/>
                    string time = Regex.Match(html, @"(?<=\btime="")[^""]*").Value;
                    double milliseconds = Convert.ToInt64(time) / 1000.0;
                    dateTime = new DateTime(1970, 1, 1).AddMilliseconds(milliseconds).ToLocalTime();
                }
            }
            catch
            {

            }
            return dateTime;
        }

        public static bool CheckForInternetConnection()
        {
            try
            {
                using (var client = new WebClient())
                using (var stream = client.OpenRead("http://www.google.com"))
                {
                    return true;
                }
            }
            catch
            {
                return false;
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            _stop();
            splitContainer1.Focus();
        }

        private void логshowhideToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (splitContainer1.Panel2Collapsed == false)
            {
                splitContainer1.Panel2Collapsed = true;
                splitContainer1.Panel2.Hide();
            }
            else
            {
                splitContainer1.Panel2Collapsed = false;
                splitContainer1.Panel2.Show();
            }
        }

        private void входToolStripMenuItem_Click(object sender, EventArgs e)
        {
            check_login();
        }

        private void выходToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void обновитьToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            // webControl1.Source = new Uri(ps.url);
            webBrowser1.Navigate(ps.url);
        }

        public void SetText(string text)
        {
         
            if (this.textBox1.InvokeRequired)
            {
                SetTextCallback d = new SetTextCallback(SetText);
                this.Invoke(d, new object[] { text });
            }
            else
            {
                this.textBox1.Text = textBox1.Text + Environment.NewLine + DateTime.Now.ToShortTimeString() + " - "+text;
                statusbar(text);  
            }
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {
            textBox1.SelectionStart = textBox1.Text.Length;
            textBox1.ScrollToCaret();
        }

        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            SetText("Подготовка к завершению");
            toolStripStatusLabel1.Text = "Завершение";

            if (xlApp !=null)
            {
                xlApp.DisplayAlerts = false;
                 if (xlWorkbook != null)  xlWorkbook.Save();

                Excel_close();

            }
            SetText("Обработка файла завершена!");
        }

        private void tESTToolStripMenuItem_Click(object sender, EventArgs e)
        {

            Double w1 = ps.w1, w2 = ps.w2, w3 = ps.w3;

            string
                    iin = "060124600163",
                    date1 = "05.01.2017",
                    vrach = "КИДИРБАЕВА ДИНАРА ТУЛЕГЕНОВНА",
                    dateNaprav = "05.01.2017",
                    code_uslugi = "B06.296.006, B06.414.006, B06.444.006, B06.449.006, B06.415.006, B06.413.006, B06.470.006, B06.442.006, B06.276.006, B06.556.006, B06.518.006, B06.195.005, B06.202.006, B06.512.006, B06.500.006, B06.123.006, B06.361.006, B06.433.006, B06.432.006, B06.527.006",
                    zaveduw = "КИДИРБАЕВА ДИНАРА ТУЛЕГЕНОВНА",
                    dateAccept = "05.01.2017",
                    vrachNapr = "КИДИРБАЕВА ДИНАРА ТУЛЕГЕНОВНА";

                int medorg = Convert.ToInt32(ps.medorg.ToString()),
                     mestApp = Convert.ToInt32(ps.medotd.ToString());



        }

 

        private void test2ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            string dateNaprav = "05.01.2017";
            //Вставляем дату
            HtmlElement date_napr = findByAttribute("input", "iid", "dt_sent");
            if (date_napr != null)
            {
                date_napr.Focus();
                simulateDigitKey(dateNaprav);
            }
            else SetText("Not found");


        }



    }
}
