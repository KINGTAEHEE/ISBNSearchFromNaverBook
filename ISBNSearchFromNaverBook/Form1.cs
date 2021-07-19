using HtmlAgilityPack;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ISBNSearchFromNaverBook
{
    public partial class Form1 : Form
    {
        List<string> isbnList = new List<string>(); // ISBN 정보를 담을 string 타입인 동적 배열
        public static DataTable dt = new DataTable(); // 결과를 담을 DataTable
        DataColumn numberCol = new DataColumn("순번", typeof(int));
        DataColumn ypCol = new DataColumn("영풍문고", typeof(string));
        DataColumn itCol = new DataColumn("인터파크", typeof(string));
        DataColumn kbCol = new DataColumn("교보문고", typeof(string));
        DataColumn yesCol = new DataColumn("예스24", typeof(string));
        DataColumn alCol = new DataColumn("알라딘", typeof(string));
        DataColumn isbnCol = new DataColumn("ISBN", typeof(string));
        DataColumn titleCol = new DataColumn("제목", typeof(string));
        DataColumn writerCol = new DataColumn("저자", typeof(string));
        DataColumn publisherCol = new DataColumn("출판사", typeof(string));
        DataColumn dateCol = new DataColumn("출간일", typeof(string));
        DataColumn pageCol = new DataColumn("페이지수", typeof(string));
        string openFile = string.Empty;

        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (ShowFileOpenDialog())
            {
                if (ISBNSearchFromNaverBook())
                {
                    Form2 form2 = new Form2(dt);
                    form2.Show();
                    this.Hide();
                }
                else
                {
                    Application.Exit();
                }
            }
        }

        public bool ISBNSearchFromNaverBook()
        {
            try
            {
                button1.Enabled = false;

                StreamReader sr = new StreamReader(openFile);

                // 파일로부터 검색하고자하는 ISBN 정보를 불러온다
                while (sr.EndOfStream == false)
                {
                    isbnList.Add(sr.ReadLine());
                }

                sr.Close();

                // 진행바 설정
                float progressBarStep = 100 / (float)isbnList.Count;

                // DataTable Column 설정
                dt.Columns.Add(numberCol);
                dt.Columns.Add(ypCol);
                dt.Columns.Add(itCol);
                dt.Columns.Add(kbCol);
                dt.Columns.Add(yesCol);
                dt.Columns.Add(alCol);
                dt.Columns.Add(isbnCol);
                dt.Columns.Add(titleCol);
                dt.Columns.Add(writerCol);
                dt.Columns.Add(publisherCol);
                dt.Columns.Add(dateCol);
                dt.Columns.Add(pageCol);

                // 네이버책에서 ISBN 검색 후 판매 목록에 영풍문고 있는지 확인하고 결과를 DataTable에 저장
                for (int i = 0; i < isbnList.Count; i++)
                {
                    // DataTable에 넣을 DataRow 설정
                    DataRow row = dt.NewRow();

                    // 네이버책에 ISBN 검색하고 노드 정보 읽기
                    var web = new HtmlWeb();
                    var doc = web.Load("https://book.naver.com/search/search.nhn?sm=sta_hty.book&sug=&where=nexearch&query=" + isbnList[i]);
                    var nodes = doc.DocumentNode.SelectNodes("//*[@id='searchBiblioList']/li/div/div/a");
                    string src = string.Empty;

                    // 네이버책 검색결과에서 노드 제대로 찾지 못할 경우 네이버 통합검색을 통해 정보 읽는다
                    if (nodes == null)
                    {
                        doc = web.Load("https://search.naver.com/search.naver?where=nexearch&sm=top_hty&fbm=1&ie=utf8&query=" + isbnList[i]);
                        nodes = doc.DocumentNode.SelectNodes("//*[@id='book-collection']/div/ul/li/div/div/a");
                    }

                    // 네이버 통합검색 페이지에서 책 노드를 제대로 못찾은 경우 두번째 시도
                    if (nodes == null)
                    {
                        nodes = doc.DocumentNode.SelectNodes("//*[@id='book-collection']/div/div[2]/div[1]/div[1]/a");
                    }

                    // 검색결과 책이 없으면 넘긴다
                    if (nodes == null)
                    {
                        row[numberCol] = i + 1;
                        row[ypCol] = "※ 네이버책에 등록되지 않은 ISBN - 직접 확인 필요";
                        row[itCol] = "";
                        row[kbCol] = "";
                        row[yesCol] = "";
                        row[alCol] = "";
                        row[isbnCol] = isbnList[i];
                        row[titleCol] = "";
                        row[writerCol] = "";
                        row[publisherCol] = "";
                        row[dateCol] = "";
                        row[pageCol] = "";
                        dt.Rows.Add(row);

                        label2.Text = (i + 1).ToString() + " / " + isbnList.Count.ToString();
                        Application.DoEvents();
                        progressBar1.Value = (int)((i + 1) * progressBarStep); // 진행바 갱신
                        continue;
                    }

                    // 검색결과 책이 있으면 노드 정보중 이미지 주소를 통해 네이버책 정보 페이지 주소를 알아낸다
                    foreach (var node in nodes)
                    {
                        string temp = node.OuterHtml;
                        temp = node.OuterHtml.Substring(node.OuterHtml.IndexOf('?') + 5, 8);
                        // 8자리가 아니고 7자리면 다시 자른다
                        if (temp.Contains("\""))
                        {
                            temp = temp.Substring(0, 7);
                        }
                        src = temp;
                        break;
                    }

                    // 네이버책 정보 페이지로 이동하고 노드 정보 읽기
                    doc = web.Load("https://book.naver.com/bookdb/book_detail.nhn?bid=" + src);
                    var buyList = doc.DocumentNode.SelectNodes("//*[@id='productListLayer']/ul"); // 바로구매 영역
                    var bookTitle = doc.DocumentNode.SelectNodes("//*[@id='container']/div[4]/div[1]/h2/a"); // 책 제목
                    var bookInfo1 = doc.DocumentNode.SelectNodes("//*[@id='container']/div[4]/div[1]/div[2]/div[2]"); // 책 정보(저자, 출판사, 출간일)
                    var bookInfo2 = doc.DocumentNode.SelectNodes("//*[@id='container']/div[4]/div[1]/div[2]/div[3]"); // 책 정보(페이지수)

                    // 노드 정보중 하나라도 null 값이 있으면 오류로 판단하고 넘긴다
                    if (buyList == null || bookTitle == null || bookInfo1 == null || bookInfo2 == null)
                    {
                        row[numberCol] = i + 1;
                        row[ypCol] = "※ 네이버책 페이지 추적 불가 - 직접 확인 필요";
                        row[itCol] = "";
                        row[kbCol] = "";
                        row[yesCol] = "";
                        row[alCol] = "";
                        row[isbnCol] = isbnList[i];
                        row[titleCol] = "";
                        row[writerCol] = "";
                        row[publisherCol] = "";
                        row[dateCol] = "";
                        row[pageCol] = "";
                        dt.Rows.Add(row);

                        label2.Text = (i + 1).ToString() + " / " + isbnList.Count.ToString();
                        Application.DoEvents();
                        progressBar1.Value = (int)((i + 1) * progressBarStep); // 진행바 갱신
                        continue;
                    }

                    // 노드 정보중에 바로구매 영역에 영풍문고가 있는지 확인하고 결과와 책 정보를 DataTable에 저장
                    foreach (var list in buyList) // 바로구매 영역
                    {
                        if (list.InnerText.Contains("영풍문고"))
                        {
                            row[numberCol] = i + 1;
                            row[ypCol] = "O";
                            row[isbnCol] = isbnList[i];
                        }
                        else
                        {
                            row[numberCol] = i + 1;
                            row[ypCol] = "X";
                            row[isbnCol] = isbnList[i];
                        }

                        if (list.InnerText.Contains("인터파크"))
                        {
                            row[numberCol] = i + 1;
                            row[itCol] = "O";
                            row[isbnCol] = isbnList[i];
                        }
                        else
                        {
                            row[numberCol] = i + 1;
                            row[itCol] = "X";
                            row[isbnCol] = isbnList[i];
                        }

                        if (list.InnerText.Contains("교보문고"))
                        {
                            row[numberCol] = i + 1;
                            row[kbCol] = "O";
                            row[isbnCol] = isbnList[i];
                        }
                        else
                        {
                            row[numberCol] = i + 1;
                            row[kbCol] = "X";
                            row[isbnCol] = isbnList[i];
                        }

                        if (list.InnerText.Contains("예스24"))
                        {
                            row[numberCol] = i + 1;
                            row[yesCol] = "O";
                            row[isbnCol] = isbnList[i];
                        }
                        else
                        {
                            row[numberCol] = i + 1;
                            row[yesCol] = "X";
                            row[isbnCol] = isbnList[i];
                        }

                        if (list.InnerText.Contains("알라딘"))
                        {
                            row[numberCol] = i + 1;
                            row[alCol] = "O";
                            row[isbnCol] = isbnList[i];
                        }
                        else
                        {
                            row[numberCol] = i + 1;
                            row[alCol] = "X";
                            row[isbnCol] = isbnList[i];
                        }


                        foreach (var title in bookTitle) // 책 제목
                        {
                            row[titleCol] = title.InnerText.Replace("&nbsp;", " ");
                        }

                        foreach (var info1 in bookInfo1) // 책 정보(저자, 출판사, 출간일)
                        {
                            string[] temp = new string[5];
                            temp = info1.InnerText.Split('|');

                            if (temp[0].Contains("글 ") || temp[1].Contains("역자 ") || temp[1].Contains("편집 ")) // 저자 정보가 파트별로 나뉘어 적혀있는 경우
                            {
                                if (temp[0].Contains("글 ") && temp[1].Contains("그림 ") && temp[2].Contains("역자 ")) // 글/그림/역자 3가지로 적혀있는 경우
                                {
                                    row[writerCol] = "글 " + (temp[0].Trim()).Substring(2).Replace("&nbsp;", " ")
                                                    + ", 그림 " + (temp[1].Trim()).Substring(3).Replace("&nbsp;", " ")
                                                    + ", 역자 " + (temp[2].Trim()).Substring(3).Replace("&nbsp;", " ");
                                    row[publisherCol] = temp[3].Trim().Replace("&nbsp;", " "); // 출판사
                                    row[dateCol] = temp[4].Trim().Replace("&nbsp;", " "); // 출간일
                                }
                                else // 글/그림 or 저자/역자 2가지로 적혀있는 경우
                                {
                                    if (temp[0].Contains("글 ") && temp[1].Contains("그림 ")) // 글/그림 2가지로 적혀있는 경우
                                    {
                                        row[writerCol] = "글 " + (temp[0].Trim()).Substring(2).Replace("&nbsp;", " ")
                                                        + ", 그림 " + (temp[1].Trim()).Substring(3).Replace("&nbsp;", " ");
                                        row[publisherCol] = temp[2].Trim().Replace("&nbsp;", " "); // 출판사
                                        row[dateCol] = temp[3].Trim().Replace("&nbsp;", " "); // 출간일
                                    }
                                    else if (temp[0].Contains("저자 ") && temp[1].Contains("역자 ")) // 저자/역자 2가지로 적혀있는 경우
                                    {
                                        row[writerCol] = "저자 " + (temp[0].Trim()).Substring(2).Replace("&nbsp;", " ")
                                                        + ", 역자 " + (temp[1].Trim()).Substring(3).Replace("&nbsp;", " ");
                                        row[publisherCol] = temp[2].Trim().Replace("&nbsp;", " "); // 출판사
                                        row[dateCol] = temp[3].Trim().Replace("&nbsp;", " "); // 출간일
                                    }
                                    else // 저자/편집 2가지로 적혀있는 경우
                                    {
                                        row[writerCol] = "저자 " + (temp[0].Trim()).Substring(2).Replace("&nbsp;", " ")
                                                        + ", 편집 " + (temp[1].Trim()).Substring(3).Replace("&nbsp;", " ");
                                        row[publisherCol] = temp[2].Trim().Replace("&nbsp;", " "); // 출판사
                                        row[dateCol] = temp[3].Trim().Replace("&nbsp;", " "); // 출간일
                                    }
                                }
                            }
                            else // 단순히 저자 하나로만 적혀있는 경우
                            {
                                row[writerCol] = (temp[0].Trim()).Substring(3).Replace("&nbsp;", " "); // 저자
                                row[publisherCol] = temp[1].Trim().Replace("&nbsp;", " "); // 출판사
                                row[dateCol] = temp[2].Trim().Replace("&nbsp;", " "); // 출간일
                            }
                        }

                        foreach (var info2 in bookInfo2) // 책 정보(페이지수)
                        {
                            string[] temp = new string[5];
                            temp = info2.InnerText.Split('|');
                            if (temp[0].Substring(temp[0].IndexOf('지') + 2).Length > 4)
                            {
                                row[pageCol] = "";
                            }
                            else
                            {
                                row[pageCol] = (temp[0].Trim()).Substring(4);
                            }
                        }
                    }

                    dt.Rows.Add(row); // 정상처리된 Row를 DataTable에 추가

                    label2.Text = (i + 1).ToString() + " / " + isbnList.Count.ToString();
                    Application.DoEvents();
                    progressBar1.Value = (int)((i + 1) * progressBarStep); // 진행바 갱신
                }

                button1.Enabled = true;
                return true;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message + "\n\n!! 오류로 인해 자동 종료 !!");
                return false;
            }
        }

        public bool ShowFileOpenDialog()
        {
            // 파일 오픈창 생성 및 설정
            OpenFileDialog ofd = new OpenFileDialog();
            ofd.Filter = "데이터 파일 (*.dat)|*.dat|텍스트 파일 (*.txt)|*.txt";

            DialogResult dr = ofd.ShowDialog();

            if (dr == DialogResult.OK)
            {
                openFile = ofd.FileName;
                return true;
            }
            else
            {
                return false;
            }
        }
    }
}
