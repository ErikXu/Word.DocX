using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using Novacode;

namespace Word.Export
{
    public class Exporter
    {
        private readonly static string TemplateFolder = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "templates");

        public static void Export(Info info, string filePath)
        {
            using (var fs = new FileStream(filePath, FileMode.Create))
            {
                var doc = Export(info);
                doc.SaveAs(fs);
            }
        }

        public static DocX Export(Info info)
        {
            //Load an empty doc and remove the first paragraph
            var filePath = Path.Combine(TemplateFolder, "Empty.docx");
            var doc = DocX.Load(filePath);
            doc.RemoveParagraphAt(0);

            var docs = new List<DocX>();

            var enterDocs = FillEnters(info.Basic, info.Enters);
            docs.AddRange(enterDocs);

            var departureDocs = FillDepartures(info.Basic, info.Departures);
            docs.AddRange(departureDocs);

            foreach (var d in docs)
            {
                doc.InsertDocument(d);
            }

            return doc;
        }

        private static IEnumerable<DocX> FillEnters(Basic basic, List<Enter> enters)
        {
            const int rowCountOfPage = 8;

            var pages = GetPages(enters.Count, rowCountOfPage);

            var docs = new List<DocX>();

            for (var i = 0; i < pages; i++)
            {
                var entersOfPage = enters.Skip(rowCountOfPage * i).Take(rowCountOfPage);

                var doc = FillEnter(basic, entersOfPage);
                docs.Add(doc);
            }

            return docs;
        }

        private static DocX FillEnter(Basic basic, IEnumerable<Enter> enters)
        {
            var templatePath = Path.Combine(TemplateFolder, "Enter.docx");

            var doc = DocX.Load(templatePath);

            doc.Paragraphs[1].ReplaceText("[项目名称]", basic.Project);
            doc.Paragraphs[3].ReplaceText("[公司名称]", basic.Company);
            doc.Paragraphs[3].ReplaceText("[班组名称]", basic.Team);
            doc.Paragraphs[3].ReplaceText("[始年]", basic.StartDate.Year.ToString(CultureInfo.InvariantCulture));
            doc.Paragraphs[3].ReplaceText("[始月]", basic.StartDate.Month.ToString(CultureInfo.InvariantCulture));
            doc.Paragraphs[3].ReplaceText("[始日]", basic.StartDate.Day.ToString(CultureInfo.InvariantCulture));
            doc.Paragraphs[3].ReplaceText("[结年]", basic.EndDate.Year.ToString(CultureInfo.InvariantCulture));
            doc.Paragraphs[3].ReplaceText("[结月]", basic.EndDate.Month.ToString(CultureInfo.InvariantCulture));
            doc.Paragraphs[3].ReplaceText("[结日]", basic.EndDate.Day.ToString(CultureInfo.InvariantCulture));
            doc.Paragraphs[4].ReplaceText("[进场数]", basic.EnterCount.ToString(CultureInfo.InvariantCulture));
            doc.Paragraphs[4].ReplaceText("[离场数]", basic.DepartureCount.ToString(CultureInfo.InvariantCulture));
            doc.Paragraphs[4].ReplaceText("[现场数]", basic.CurrentCount.ToString(CultureInfo.InvariantCulture));

            if (enters != null)
            {
                var table = doc.Tables[0];

                var index = 2;

                foreach (var item in enters)
                {
                    var row = table.Rows[index];

                    row.Cells[0].Paragraphs[0].Append(item.Order.ToString(CultureInfo.InvariantCulture));
                    row.Cells[1].Paragraphs[0].Append(item.Name);
                    row.Cells[2].Paragraphs[0].Append(item.IdNum);
                    row.Cells[3].Paragraphs[0].Append(item.EnterDate.ToString("yyyy-MM-dd"));
                    row.Cells[4].Paragraphs[0].Append(item.Job);
                    row.Cells[5].Paragraphs[0].Append(item.IsRecord ? "是" : "否");
                    row.Cells[6].Paragraphs[0].Append(item.HasContract ? "是" : "否");

                    index++;
                }
            }

            return doc;
        }

        private static IEnumerable<DocX> FillDepartures(Basic basic, List<Departure> departures)
        {
            const int rowCountOfPage = 9;

            var pages = GetPages(departures.Count, rowCountOfPage);

            var docs = new List<DocX>();

            for (var i = 0; i < pages; i++)
            {
                var departuresOfPage = departures.Skip(rowCountOfPage * i).Take(rowCountOfPage);

                var doc = FillDeparture(basic, departuresOfPage);
                docs.Add(doc);
            }

            return docs;
        }

        private static DocX FillDeparture(Basic basic, IEnumerable<Departure> departures)
        {
            var templatePath = Path.Combine(TemplateFolder, "Departure.docx");

            var doc = DocX.Load(templatePath);

            doc.Paragraphs[1].ReplaceText("[项目名称]", basic.Project);
            doc.Paragraphs[3].ReplaceText("[公司名称]", basic.Company);
            doc.Paragraphs[3].ReplaceText("[班组名称]", basic.Team);
            doc.Paragraphs[3].ReplaceText("[始年]", basic.StartDate.Year.ToString(CultureInfo.InvariantCulture));
            doc.Paragraphs[3].ReplaceText("[始月]", basic.StartDate.Month.ToString(CultureInfo.InvariantCulture));
            doc.Paragraphs[3].ReplaceText("[始日]", basic.StartDate.Day.ToString(CultureInfo.InvariantCulture));
            doc.Paragraphs[3].ReplaceText("[结年]", basic.EndDate.Year.ToString(CultureInfo.InvariantCulture));
            doc.Paragraphs[3].ReplaceText("[结月]", basic.EndDate.Month.ToString(CultureInfo.InvariantCulture));
            doc.Paragraphs[3].ReplaceText("[结日]", basic.EndDate.Day.ToString(CultureInfo.InvariantCulture));

            if (departures != null)
            {
                var table = doc.Tables[0];

                var index = 2;

                foreach (var item in departures)
                {
                    var row = table.Rows[index];

                    row.Cells[0].Paragraphs[0].Append(item.Order.ToString(CultureInfo.InvariantCulture));
                    row.Cells[1].Paragraphs[0].Append(item.Name);
                    row.Cells[2].Paragraphs[0].Append(item.IdNum);
                    row.Cells[3].Paragraphs[0].Append(item.DepartureDate.ToString("yyyy-MM-dd"));
                    row.Cells[4].Paragraphs[0].Append(item.WorkDays.ToString(CultureInfo.InvariantCulture) + "天");
                    row.Cells[5].Paragraphs[0].Append(item.Job);
                    row.Cells[6].Paragraphs[0].Append(item.IsSettled ? "已结算" : "未结算");
                    row.Cells[7].Paragraphs[0].Append(item.IsCommited ? "是" : "否");

                    index++;
                }
            }

            return doc;
        }

        private static int GetPages(int totalCount, int countOfPage)
        {
            var pages = totalCount / countOfPage;

            if (pages == 0)
            {
                pages++;
            }
            else
            {
                var mod = totalCount % countOfPage;

                if (mod > 0)
                {
                    pages++;
                }
            }

            return pages;
        }
    }
}
