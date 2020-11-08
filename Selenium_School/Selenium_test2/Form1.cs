using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Selenium_School
{
    public partial class Form1 : Form
    {
        Boolean mjc_Flag = false;
        int LimitClass = 0;
        BackgroundWorker bw1 = new BackgroundWorker();
        BackgroundWorker bw2 = new BackgroundWorker();
        BackgroundWorker bw3 = new BackgroundWorker();
        BackgroundWorker bw4 = new BackgroundWorker();

        List<int> ClassMinArr = new List<int>();
        List<int> ClassMaxArr = new List<int>();
        List<String> EclassLink = new List<String>();

        ChromeDriver driver;
        String[] subjectName;
        String[] subjectLink;
        bool loginFlag = false;
        int EclassLinkCount;

        String nowhandle;  //현재 창의 핸들찾기
        IList<String> handles;  //모든 창의 핸들 찾기

        public Form1()
        {
            InitializeComponent();
            bw1.DoWork += new DoWorkEventHandler(backgroundWorker1_DoWork);
            bw2.DoWork += new DoWorkEventHandler(backgroundWorker2_DoWork);
            bw3.DoWork += new DoWorkEventHandler(backgroundWorker3_DoWork);
            bw4.DoWork += new DoWorkEventHandler(backgroundWorker4_DoWork);



            listView1.View = View.Details;
            listView1.BeginUpdate();
            textBox1.Text = "2019671062";
            textBox2.Text = "P@ssw0rdadds";
        }
        private void ExecuteScript(IJavaScriptExecutor driver, string script)
        {
            IJavaScriptExecutor js = (IJavaScriptExecutor)driver;
            js.ExecuteScript(script);
        }

        //비동기적 실행 (driver, 자바스크립트)
        private void ExecuteAsyncScript(IJavaScriptExecutor driver, string script)
        {
            IJavaScriptExecutor js = (IJavaScriptExecutor)driver;
            js.ExecuteAsyncScript(script);
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (!bw2.IsBusy)
            {
                bw2.RunWorkerAsync();
                return;
            }
            MessageBox.Show("이미 실행 중 입니다.");
        }

        private void backgroundWorker1_DoWork(object sender, DoWorkEventArgs e)
        {
            ChromeOptions options = new ChromeOptions();
            ChromeDriverService driverService = ChromeDriverService.CreateDefaultService();
            driverService.HideCommandPromptWindow = true;

            driver = new ChromeDriver(driverService, options);

            driver.Url = "chrome://settings/content/siteDetails?site=http%3A%2F%2Fcyber.mjc.ac.kr";
            Thread.Sleep(2000);
            for (int idx = 0; idx < 7; idx++)
            {
                driver.FindElementByCssSelector("body").SendKeys(OpenQA.Selenium.Keys.Tab);
            }
            driver.FindElementByCssSelector("body > settings-ui").SendKeys(OpenQA.Selenium.Keys.Down);

            driver.FindElementByCssSelector("body").SendKeys(OpenQA.Selenium.Keys.Tab);
            driver.FindElementByCssSelector("body").SendKeys(OpenQA.Selenium.Keys.Tab);
            driver.FindElementByCssSelector("body > settings-ui").SendKeys(OpenQA.Selenium.Keys.Down);

            Thread.Sleep(300);
            driver.Url = "https://member.mjc.ac.kr/member/login.do?returnUrl=http://cyber.mjc.ac.kr/index.jsp";
            button1.Enabled = true;
        }

        private void backgroundWorker2_DoWork(object sender, DoWorkEventArgs e)
        {
            IJavaScriptExecutor js = (IJavaScriptExecutor)driver;
            driver.Url = "https://member.mjc.ac.kr/member/login.do?returnUrl=http://cyber.mjc.ac.kr/index.jsp";
            driver.FindElementById("userId").SendKeys(textBox1.Text);
            Thread.Sleep(500);
            driver.FindElementById("userPw").SendKeys(textBox2.Text);
            Thread.Sleep(500);
            driver.FindElementByClassName("btn_blue").Click();
            Thread.Sleep(500);

            try
            {
                loginFlag = false;
                driver.SwitchTo().Alert().Accept();
                //비밀번호 틀렸을 경우


            }
            catch
            {
                Thread.Sleep(1000);
                loginFlag = true;
                button1.Enabled = false;
                //로그인 성공시
                driver.Navigate().Refresh();
                //driver.FindElement(By.CssSelector("#intro > div > div > div > a")).Click();
                Thread.Sleep(2500);

                js.ExecuteScript("document.querySelector(\"frame\").contentWindow.document.querySelector(\"#id\").value = \"" + textBox1.Text + "\";");
                js.ExecuteScript("document.querySelector(\"frame\").contentWindow.document.querySelector(\"#pw\").value = \"" + textBox2.Text + "\";");
                js.ExecuteScript("document.querySelector(\"frame\").contentWindow.document.querySelector(\"#loginForm-member > fieldset > p:nth-child(3) > a\").click();");

                Thread.Sleep(3000);
                var ulLength_var = js.ExecuteScript(
                   "return document.querySelector(\"frame\").contentWindow.document.querySelector(\"frame\").contentWindow.document.querySelector(\"#mCSB_2_container\").querySelectorAll(\"span.boardTxt\").length"
                );
                int ulLength = Convert.ToInt32(ulLength_var);
                subjectName = new string[ulLength];
                subjectLink = new string[ulLength];

                try
                {
                    for (int idx = 0; idx < ulLength; idx++)
                    {
                        var ulName = js.ExecuteScript(
                            "return document.querySelector(\"frame\").contentWindow.document.querySelector(\"frame\").contentWindow.document.querySelector(\"#mCSB_2_container\").querySelectorAll(\"span.boardTxt\")[" + idx + "].textContent;"
                        );
                        var ulLink = js.ExecuteScript(
                            "return document.querySelector(\"frame\").contentWindow.document.querySelector(\"frame\").contentWindow.document.querySelector(\"#mCSB_2_container\").querySelectorAll(\"button\")[" + idx + "].getAttribute(\"onclick\");"
                        );
                        subjectName[idx] = (String)ulName;
                        subjectLink[idx] = (String)ulLink;

                        listView1.Items.Add(subjectName[idx]);

                    }
                    button2.Enabled = true;

                }
                catch
                {
                    MessageBox.Show("Error");
                }

            }
        }
        private void TapChange(ChromeDriver driver, String handle)
        {
            driver.SwitchTo().Window(handle);
        }
        private void Form1_Load(object sender, EventArgs e)
        {
            button1.Enabled = false;
            button2.Enabled = false;
            button3.Enabled = false;

            bw1.RunWorkerAsync();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            button2.Enabled = false;
            if (!bw3.IsBusy)
                bw3.RunWorkerAsync();
            else
                MessageBox.Show("이미 실행 중 입니다.");
        }

        private void button3_Click(object sender, EventArgs e)
        {
            if (!bw4.IsBusy)
                bw4.RunWorkerAsync();
            else
                MessageBox.Show("이미 실행 중 입니다.");
        }

        private void backgroundWorker3_DoWork(object sender, DoWorkEventArgs e)
        {
            // 창크기 값이 없는 경우
            IJavaScriptExecutor js = (IJavaScriptExecutor)driver;
            int index = 0;
            //ExecuteAsyncScript(driver, "return t = document.querySelector(\"frame\").contentWindow.document.querySelector(\"#mCSB_2_container > ul > li:nth-child(1) > a > span.boardBtn > button\"); return t;");

            //var return_value = js.ExecuteScript("return document.querySelector(\"frame\").contentWindow.document.querySelector(\"#mCSB_2_container > ul > li:nth-child(1) > a > span.boardBtn > button\").parentNode.parentNode.parentNode.parentNode.innerText;");
            //MessageBox.Show(return_value.ToString());
            Thread.Sleep(500);

            var courseId = new object();
            try
            {
                courseId = js.ExecuteScript("return document.querySelector(\"frame\").contentWindow.document.querySelector(\"frame\").contentWindow.document.querySelector(\"#headerContent > h1 > a\").attributes.href.textContent.split(\"courseId=\")[1];");
            }
            catch
            {
                button2.Enabled = true;
                return;
            }

            js.ExecuteScript(
                "var script = document.createElement('script');\n" +
                "script.src = 'http://cyber.mjc.ac.kr/lmsdata/js/jquery-1.11.1.min.js';\n" +
                "script.type = 'text/javascript';\n" +
                "document.getElementsByTagName('head')[0].appendChild(script);\n" +

                "customData = document.querySelector(\"frame\").contentWindow.document.querySelector(\"frame\").contentWindow.customData\n" +
                "API_1484_11 = document.querySelector(\"frame\").contentWindow.document.querySelector(\"frame\").contentWindow.API_1484_11;\n" +
                "window.selectedLessonItemId = document.querySelector(\"frame\").contentWindow.document.querySelector(\"frame\").contentWindow.selectedLessonItemId;\n" +
                "window.studyWindow = document.querySelector(\"frame\").contentWindow.document.querySelector(\"frame\").contentWindow.studyWindow;\n" +
                "window.learningControl = document.querySelector(\"frame\").contentWindow.document.querySelector(\"frame\").contentWindow.learningControl;\n" +
                "window.befContStudyYn = document.querySelector(\"frame\").contentWindow.document.querySelector(\"frame\").contentWindow.befContStudyYn;\n" +
                "window.curContStudyYn = document.querySelector(\"frame\").contentWindow.document.querySelector(\"frame\").contentWindow.curContStudyYn;\n" +
                "window.C4Enable = document.querySelector(\"frame\").contentWindow.document.querySelector(\"frame\").contentWindow.C4Enable;\n" +
                "window.C5Enable = document.querySelector(\"frame\").contentWindow.document.querySelector(\"frame\").contentWindow.C5Enable;\n" +
                "window.attendYn = document.querySelector(\"frame\").contentWindow.document.querySelector(\"frame\").contentWindow.attendYn;\n" +
                "window.curLessonElementId = document.querySelector(\"frame\").contentWindow.document.querySelector(\"frame\").contentWindow.curLessonElementId;\n" +
                "window.curScoId = document.querySelector(\"frame\").contentWindow.document.querySelector(\"frame\").contentWindow.curScoId;\n" +


                "window.lessonObject = new Object();\n" +
                "window.courseDTO = new Object();\n"
            );

            js.ExecuteScript(
                "window.test = function() {" +
                    "alert(\"123\")" +
                "}\n" +
                "window.certificationCompleted = function() {\n" +

                    "var f = document.courseForm;\n" +
                    "var lessonElementId = lessonObject.lessonElementId;\n" +
                    "var lessonContentsId = lessonObject.lessonContentsId;\n" +
                    "var windowWidth = lessonObject.windowWidth;\n" +
                    "var windowHeight = lessonObject.windowHeight;\n" +

                    "var url = \"/Lesson.do?cmd=viewStudyContentsForm&studyRecordDTO.lessonElementId=\" + lessonElementId + \"&studyRecordDTO.lessonContentsId=\" + lessonContentsId + \"&courseDTO.courseId=" + courseId.ToString() + "\";\n" +
                    "url += \"&menuId=menu_00099\";\n" +

                    "url += \"&pus=N&pcr=N&pre=N\";\n" +

                    "var winWidth = windowWidth;\n" +
                    "var winHeight = windowHeight;\n" +
                    "if (winWidth == \"0\" || winWidth == \"\")\n" +
                        "{\n" +
                            "winWidth = 800;\n" +
                        "}\n" +
                        "if (winHeight == \"0\" || winHeight == \"\")\n" +
                        "{\n" +
                            "winHeight = 600;\n" +
                        "}\n" +
                        "console.log(\"2\");\n" +
                        "var winHeight_plus = eval(winHeight) + 45;\n" +

                        "console.log(\"3\");\n" +
                        "var top = ((screen.availHeight - winHeight) / 2 + 40);\n" +
                        "var left = ((screen.availWidth - winWidth) / 2);\n" +

                        "console.log(\"4\");\n" +
                        "studyWindow = window.open(url);\n" +
                        "studyWindow.open = true\n" +
                    "}"
        );
            js.ExecuteScript(
                    "window.viewStudyContents = function(lessonElementId, lessonContentsId, windowWidth, windowHeight, learningControl, lessonCnt){" +

                    "console.log(\"Enter\");" +
                    "window.f = document.courseForm;\n" +
                    "lessonCnt = 0;\n" +
                    "lessonObject.lessonElementId = lessonElementId;\n" +
                    "lessonObject.lessonContentsId = lessonContentsId;\n" +
                    "lessonObject.windowWidth = windowWidth;\n" +
                    "lessonObject.windowHeight = windowHeight;\n\n" +

                    "courseDTO.courseId = \"" + courseId.ToString() + "\"\n" +
                    "certificationCompleted();\n" +
                "}\n\n"
            );

            var EclassLinkCountScript = js.ExecuteScript(
                "return document.querySelector(\"frame\").contentWindow.document.querySelector(\"frame\").contentWindow.document.querySelectorAll(\".btn-orange\").length;"
            );
            EclassLinkCount = Convert.ToInt32(EclassLinkCountScript);
            if (EclassLinkCount == 0)
            {
                button2.Enabled = true;
                return;
            }
            String _Message = "";

            for (int idx = 0; idx < EclassLinkCount; idx++)
            {
                var ClassMin = js.ExecuteScript(
                   "window.sp = document.querySelector(\"frame\").contentWindow.document.querySelector(\"frame\").contentWindow.document.querySelectorAll(\".btn-orange\")[" + idx + "].parentElement.parentElement.querySelector(\".bar\").textContent.trim().split(\"분\");\n" +
                   "if(sp.length == 3) {" +
                        "return 0;" +
                   "}" +
                   "return sp[0].substring(1, sp[0].length);"
                );
                var ClassMax = js.ExecuteScript(
                   "window.sp = document.querySelector(\"frame\").contentWindow.document.querySelector(\"frame\").contentWindow.document.querySelectorAll(\".btn-orange\")[" + idx + "].parentElement.parentElement.querySelector(\".bar\").textContent.trim().split(\"분\");\n" +
                   "return sp[sp.length - 2].trim().substring(2,sp[sp.length - 2].length);"
                );
                //MessageBox.Show(ClassMin + "/" + ClassMax);

                var Temp = js.ExecuteScript(
                    "return document.querySelector(\"frame\").contentWindow.document.querySelector(\"frame\").contentWindow.document.querySelectorAll(\".btn-orange\")[" + idx + "].getAttribute(\"href\");"
                );
                EclassLink.Add(Temp.ToString());
                _Message += EclassLink[idx] + "\n";

                if (Convert.ToInt32(ClassMax) != 1)
                {
                    js.ExecuteScript(EclassLink[idx]); //실행
                    Thread.Sleep(200);

                    ClassMinArr.Add(Convert.ToInt32(ClassMin));
                    ClassMaxArr.Add(Convert.ToInt32(ClassMax));

                    index++;
                }
            }

            Thread.Sleep(1000);

            button2.Enabled = true;
            button3.Enabled = true;
        }

        private void backgroundWorker4_DoWork(object sender, DoWorkEventArgs e)
        {
            IJavaScriptExecutor js;

            nowhandle = driver.CurrentWindowHandle;
            handles = driver.WindowHandles;

            for (int idx = handles.Count - 1; idx >= 0; idx--)
            {
                driver.SwitchTo().Window(handles[idx]);
                js = (IJavaScriptExecutor)driver;
                //js.ExecuteScript("document.querySelector(\"frame\").contentWindow.document.querySelector(\"frame\").contentWindow.document.querySelector(\"#player > div.fp-player > div.fp-ui > div.fp-controls > div.fp-volume > a\").click();");
                Thread.Sleep(300);
                //
            }

            nowhandle = driver.CurrentWindowHandle;
            handles = driver.WindowHandles;

            for (int idx = 1; idx < handles.Count; idx++)
            {
                driver.SwitchTo().Window(handles[idx]);
                js = (IJavaScriptExecutor)driver;

                js.ExecuteScript(
                    "window.editStudyRecord = function(finish) {\n" +
                        "is_end_flag = finish;\n" +
                        "LessonWork.editStudyRecord(StudyRecordDTO, learningControl, inStudyDate, editStudyRecordCallback);\n" +
                        "if (!finish) {\n" +
                        "videoSessionTime = 0;\n" +
                        "videoStartTime = Math.round((new Date()).getTime() / 1000);\n" +
                        "setTimeout(\"editStudyRecord(false)\", 1000);\n" +
                        "}\n" +
                    "}\n" +

                    "window.shutdown = function(time) {\n" +
                        "console.log(time + \"분 뒤에 종료됩니다.\");\n" +
                        "setTimeout(function() {\n" +
                        "window.close();\n" +
                        "}, time * 60 * 1000);\n" +
                        "}\n" +
                        "\n" +
                        "//ex) shutdown(15); 15분 뒤에 종료\n" +
                    "console.log(\"" + ClassMaxArr[handles.Count - idx - 1] + ":" + ClassMinArr[handles.Count - idx - 1] + "\");\n" +
                    "shutdown(" + (ClassMaxArr[handles.Count - idx - 1] - ClassMinArr[handles.Count - idx - 1] + 2) + ");\n"

                 );
                Thread.Sleep(1500);
            }
            MessageBox.Show("적용 완료");
        }

        private void listView1_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void listView1_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            nowhandle = driver.CurrentWindowHandle;
            handles = driver.WindowHandles;
            driver.SwitchTo().Window(handles[0]);

            String scriptTemp;
            IJavaScriptExecutor js = (IJavaScriptExecutor)driver;
            int selectIndex = 0;
            for (int idx = 0; idx < listView1.Items.Count; idx++)
            {
                if (listView1.SelectedItems[0].Text.Equals(listView1.Items[idx].Text))
                {
                    selectIndex = idx;
                    break;
                }

            }
            scriptTemp = subjectLink[selectIndex].Substring(11, subjectLink[selectIndex].Length - 24);
            String[] splitScript = scriptTemp.Split('\'');
            String[] splitResult = splitScript[1].Split(',');
            String url = "http://cyber.mjc.ac.kr/Main.do?cmd=moveCourseMenu&mainDTO.parentMenuId=menu_00091&courseDTO.courseId=" + splitResult[0] + "&courseDTO.shareCourseId=&fronGubun=A&mainDTO.menuId=menu_00099";

            if(!mjc_Flag)
            {
                String script = "document.querySelector(\"frame\").contentWindow.document.querySelector(\"frame\").contentWindow." + scriptTemp;
                js.ExecuteScript(script);
                mjc_Flag = true;

                Thread.Sleep(2000);

            } 
            js.ExecuteScript("location.replace(\"" + url + "\");");
            

           
             
        }

        private void button4_Click(object sender, EventArgs e)
        {
            String scriptTemp;
            IJavaScriptExecutor js = (IJavaScriptExecutor)driver;
            int selectIndex = 0;
            try
            {
                String temp = listView1.SelectedItems[0].Text;

                for (int idx = 0; idx < listView1.Items.Count; idx++)
                {
                    if (listView1.SelectedItems[0].Text.Equals(listView1.Items[idx].Text))
                    {
                        selectIndex = idx;
                        break;
                    }
                }
                scriptTemp = subjectLink[selectIndex].Substring(11, subjectLink[selectIndex].Length - 24);
                String script = "document.querySelector(\"frame\").contentWindow.document.querySelector(\"frame\").contentWindow." + scriptTemp;
                js.ExecuteScript(script);
            }
            catch
            {
                MessageBox.Show("목록을 선택해주세요.");
            }
        }
    }
}
