using ClosedXML.Excel;
using OpenQA.Selenium;
using OpenQA.Selenium.Edge;
using OpenQA.Selenium.Support.UI;

namespace Koya.Uehara.Create.InputiQube
{
    public class Program
    {
        delegate void RunMethod();

        public static async Task Main(string[] args)
        {
            try
            {
                switch (args[0])
                {
                    case "FromExcelToiQube":
                        await FromExcelToiQube(args);
                        break;
                    case "FromiQubeToExcel":
                        FromiQubeToExcel(args);
                        break;
                };
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex);
            }
            finally
            {
                await Task.Delay(4000);
            }
        }

        public static async Task FromExcelToiQube(string[] args)
        {
            // iQubeに入力するデータ
            var inputData = new List<RowData>();

            try
            {
                //すでに開いているタイムカード.xlsmを開けるようにする
                var readFile = args[1] + "\\タイムカード.xlsm";
                FileStream fileStream = new(readFile, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);

                Console.WriteLine("読み込み先:" + readFile);

                // ワークブックを開いて、1ページ目のワークシートを開く
                using var workbook = new XLWorkbook(fileStream);
                var worksheet = workbook.Worksheet(1);

                // エクセルから対象年月日を読み込む
                var year = (int)worksheet.Cell("B3").Value;
                var month = (int)worksheet.Cell("B4").Value;

                // 基準年月日を作成
                var fixedDate = new DateTime(year, month, 1);

                // 範囲内を読み込む
                var tableRange = worksheet.Range("B7:H37").Rows();

                foreach (var row in tableRange)
                {
                    if (!(string.IsNullOrEmpty(row.Cell(3).Value.ToString())
                        && string.IsNullOrEmpty(row.Cell(4).Value.ToString())
                        && string.IsNullOrEmpty(row.Cell(5).Value.ToString())
                        && string.IsNullOrEmpty(row.Cell(6).Value.ToString())))
                    {
                        /// 出社～戻りに全て・一部入力あり

                        // 入力がある日付の取得
                        TimeSpan inputDay;
                        try
                        {
                            var targetDate = DateTime.Parse(row.Cell(1).Value.ToString());
                            inputDay = targetDate - fixedDate;
                        }
                        catch
                        {
                            var targetDate = DateTime.FromOADate((double)row.Cell(1).Value);
                            inputDay = targetDate - fixedDate;
                        }

                        // 出社
                        string ArrHou = "-1";
                        string ArrMin = "-1";
                        if (!string.IsNullOrEmpty(row.Cell(3).Value.ToString()))
                        {
                            ArrHou = DateTime.Parse(row.Cell(3).Value.ToString()).Hour.ToString();
                            ArrMin = DateTime.Parse(row.Cell(3).Value.ToString()).Minute.ToString();
                        }

                        // 退社
                        string LeaHou = "-1";
                        string LeaMin = "-1";
                        if (!string.IsNullOrEmpty(row.Cell(4).Value.ToString()))
                        {
                            LeaHou = DateTime.Parse(row.Cell(4).Value.ToString()).Hour.ToString();
                            LeaMin = DateTime.Parse(row.Cell(4).Value.ToString()).Minute.ToString();
                        }

                        // 外出
                        string OutHou = "-1";
                        string OutMin = "-1";
                        if (!string.IsNullOrEmpty(row.Cell(5).Value.ToString()))
                        {
                            OutHou = DateTime.Parse(row.Cell(5).Value.ToString()).Hour.ToString();
                            OutMin = DateTime.Parse(row.Cell(5).Value.ToString()).Minute.ToString();
                        }

                        // 戻り
                        string RetHou = "-1";
                        string RetMin = "-1";
                        if (!string.IsNullOrEmpty(row.Cell(6).Value.ToString()))
                        {
                            RetHou = DateTime.Parse(row.Cell(6).Value.ToString()).Hour.ToString();
                            RetMin = DateTime.Parse(row.Cell(6).Value.ToString()).Minute.ToString();
                        }

                        // iQubeに入力するデータに格納する
                        inputData.Add(new RowData(inputDay.Days, ArrHou, ArrMin, LeaHou, LeaMin, OutHou, OutMin, RetHou, RetMin));
                    }
                    else
                    {
                        // 出社～戻りに全て入力なし
                        // nope
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex);
                return;
            }

            // アプリケーション(exe)と同じディレクトリにあるファイルを指定する
            string seleniumManager = AppDomain.CurrentDomain.BaseDirectory;

            // Edgeを起動
            using var driver = new EdgeDriver(seleniumManager);

            try
            {
                // 暗黙的な待機
                driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(10);

                // iQubeログインページを開く
                driver.Url = "https://app.iqube.net/login?code=kwc";

                // メールアドレスとパスワードを入力
                driver.FindElement(By.XPath("//*[@id=\"user_email\"]")).SendKeys(args[2]);
                driver.FindElement(By.XPath("//*[@id=\"user_password\"]")).SendKeys(args[3]);

                // ログイン
                driver.FindElement(By.XPath("//*[@id=\"main\"]/div/div[2]/div[2]/form")).Submit();

                // タイムカードページに遷移
                driver.Url = "https://app.iqube.net/time_cards";

                // 編集回数分ループ
                foreach (var data in inputData)
                {
                    // 編集ボタンをクリックする
                    // 読み込み例外が発生するため、例外が発生したらリトライ
                    var retryCount = 0;
                    while (true)
                    {
                        try
                        {
                            // 編集ボタンをクリック
                            var edits = driver.FindElement(By.Id("main")).FindElements(By.ClassName("ic_edit"));
                            edits.ElementAt(data.Day).Click();

                            // 出社
                            if (!data.ArrHou.Equals("-1"))
                            {
                                var selectArrHou = new SelectElement(driver.FindElement(By.Name("time_card[arrival_hour]")));
                                var selectArrMin = new SelectElement(driver.FindElement(By.XPath("//*[@id=\"time_card_arrival_minute\"]")));
                                selectArrHou.SelectByValue(data.ArrHou);
                                selectArrMin.SelectByValue(data.ArrMin);
                            }

                            // 退社
                            if (!data.LeaHou.Equals("-1"))
                            {
                                var selectLeaHou = new SelectElement(driver.FindElement(By.XPath("//*[@id=\"time_card_leaving_hour\"]")));
                                var selectLeaMin = new SelectElement(driver.FindElement(By.XPath("//*[@id=\"time_card_leaving_minute\"]")));
                                selectLeaHou.SelectByValue(data.LeaHou);
                                selectLeaMin.SelectByValue(data.LeaMin);
                            }

                            // 外出
                            if (!data.OutHou.Equals("-1"))
                            {
                                var selectOutHou = new SelectElement(driver.FindElement(By.XPath("//*[@id=\"time_card_outing_hour\"]")));
                                var selectOutMin = new SelectElement(driver.FindElement(By.XPath("//*[@id=\"time_card_outing_minute\"]")));
                                selectOutHou.SelectByValue(data.OutHou);
                                selectOutMin.SelectByValue(data.OutMin);
                            }

                            // 戻り
                            if (!data.RetHou.Equals("-1"))
                            {
                                var selectRetHou = new SelectElement(driver.FindElement(By.XPath("//*[@id=\"time_card_returning_hour\"]")));
                                var selectRetMin = new SelectElement(driver.FindElement(By.XPath("//*[@id=\"time_card_returning_minute\"]")));
                                selectRetHou.SelectByValue(data.RetHou);
                                selectRetMin.SelectByValue(data.RetMin);
                            }

                            break;
                        }
                        catch (WebDriverException)
                        {
                            // 0.5秒毎にリトライ。10回まで
                            if (retryCount == 10)
                            {
                                throw new Exception();
                            }

                            await Task.Delay(500);
                            retryCount++;
                        }
                    }

                    // 編集を閉じる
                    driver.FindElement(By.Id("submit_button")).Click();
                }
            }
            catch (IndexOutOfRangeException)
            {
                Console.WriteLine("ログインできませんでした。");
                return;
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex);
                return;
            }
            finally
            {
                driver.Close();
            }

            Console.WriteLine("正常に完了しました。");
        }

        [System.Diagnostics.CodeAnalysis.SuppressMessage("Style", "IDE0039:ローカル関数を使用します", Justification = "<保留中>")]
        public static void FromiQubeToExcel(string[] args)
        {
            string seleniumManager = AppDomain.CurrentDomain.BaseDirectory;

            // ヘッドレスモードで実行
            var options = new EdgeOptions();
            options.AddArgument("--headless");
            using var driver = new EdgeDriver(seleniumManager, options);

            // iQubeに入力するデータ
            var outputData = new List<string[]>();

            try
            {
                driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(10);

                driver.Url = "https://app.iqube.net/login?code=kwc";
                driver.FindElement(By.XPath("//*[@id=\"user_email\"]")).SendKeys(args[2]);
                driver.FindElement(By.XPath("//*[@id=\"user_password\"]")).SendKeys(args[3]);
                driver.FindElement(By.XPath("//*[@id=\"main\"]/div/div[2]/div[2]/form")).Submit();

                driver.Url = "https://app.iqube.net/time_cards";

                // 回数を取得
                var tableCount = driver.FindElement(By.Id("main")).FindElements(By.ClassName("ic_edit")).Count;

                for (var i = 0; i < tableCount; i++)
                {
                    // iQube上で「---」となっている箇所を、空白に変更する
                    Func<string, string> notEntredConfirm = str => str.Equals("---") ? string.Empty : str;

                    // 列情報を取得する
                    var Arr = notEntredConfirm(driver.FindElement(By.XPath($"//*[@id=\"contents\"]/table/tbody/tr[{2 + i}]//td[3]")).Text.Trim());
                    var Lea = notEntredConfirm(driver.FindElement(By.XPath($"//*[@id=\"contents\"]/table/tbody/tr[{2 + i}]//td[4]")).Text.Trim());
                    var Out = notEntredConfirm(driver.FindElement(By.XPath($"//*[@id=\"contents\"]/table/tbody/tr[{2 + i}]//td[5]")).Text.Trim());
                    var Ret = notEntredConfirm(driver.FindElement(By.XPath($"//*[@id=\"contents\"]/table/tbody/tr[{2 + i}]//td[6]")).Text.Trim());

                    outputData.Add(new string[] { Arr, Lea, Out, Ret });
                }
            }
            catch (IndexOutOfRangeException)
            {
                Console.WriteLine("ログインできませんでした。");
                return;
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex);
                return;
            }
            finally
            {
                driver.Close();
            }

            try
            {
                var readFile = args[1] + "\\タイムカード.xlsm";
                FileStream fileStream = new(readFile, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
                Console.WriteLine("読み込み先:" + readFile);

                using var workbook = new XLWorkbook(fileStream);
                var worksheet = workbook.Worksheet(1);

                // 既存を読み込み、書き込み先を作成する
                var writeFile = args[1] + "\\(New)タイムカード.xlsm";
                workbook.SaveAs(writeFile);

                // 書き込む新規ワークブックを開く
                var newWorkbook = new XLWorkbook(writeFile);
                var newWorksheet = newWorkbook.Worksheet(1);

                // データを一括挿入する
                newWorksheet.Cell("D7").InsertData(outputData);

                // ワークブックを保存
                Console.WriteLine("書き込み先:" + writeFile);
                newWorkbook.SaveAs(writeFile);
            }
            catch (IOException)
            {
                Console.WriteLine("書き込みに失敗しました。");
                return;
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex);
                return;
            }

            Console.WriteLine("正常に完了しました。");
        }
    }

    public record RowData
        (
            int Day,
            string ArrHou,
            string ArrMin,
            string LeaHou,
            string LeaMin,
            string OutHou,
            string OutMin,
            string RetHou,
            string RetMin
        );
}