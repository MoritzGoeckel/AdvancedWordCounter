using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Office.Interop.Word;
using System.Text.RegularExpressions;

namespace WordCounterAddin
{
    public partial class StatisticsForm : Form
    {
        private Document document;

        public StatisticsForm(Document document)
        {
            this.document = document;
            InitializeComponent();
        }

        private void StatisticsForm_Load(object sender, EventArgs e)
        {
            List<Paragraph> paragraphs = new List<Paragraph>();
            foreach (Paragraph p in document.Paragraphs)
            {
                Style style = p.get_Style() as Style;
                if (style.NameLocal.StartsWith("Heading"))
                    paragraphs.Add(p);
            }

            Regex r = new Regex("\\(([0-9]+)\\)");

            for(int i = 0; i < paragraphs.Count; i++)
            {
                int level = getLevel(paragraphs[i]);

                int subchapters = 0;
                for (int j = i + 1; j < paragraphs.Count && getLevel(paragraphs[j]) > level; j++)
                    subchapters++;

                int wordCount = 0;

                if(paragraphs.Count > i + subchapters + 1)
                    wordCount = document.Range(paragraphs[i].Range.Start, paragraphs[i + subchapters + 1].Range.Start).ComputeStatistics(WdStatistic.wdStatisticWords);
                else
                    wordCount = document.Range(paragraphs[i].Range.Start, document.Content.End).ComputeStatistics(WdStatistic.wdStatisticWords);

                int wordGoal = -1;
                foreach (Match m in r.Matches(paragraphs[i].Range.Text))
                    wordGoal = Convert.ToInt32(m.Value.Substring(1, m.Value.Length - 2));

                /*string indention = "";
                while (indention.Length < level)
                    indention += "#";*/

                int progress = (int)Math.Round((float)wordCount / (float)wordGoal * 100f);
                if (wordCount == 0)
                    progress = 0;

                string title = paragraphs[i].Range.Text;

                ListViewItem item = new ListViewItem(new[] { level.ToString(), title, subchapters != 0 ? subchapters.ToString() : "", wordCount.ToString(), wordGoal != -1 ? wordGoal.ToString() : "", wordGoal != -1 ? progress + "%" : "" });

                if ((wordGoal != -1 && progress < 30) || title.Contains("(todo)"))
                {
                    item.BackColor = Color.DarkRed;
                    item.ForeColor = Color.White;
                }
                if ((wordGoal != -1 && progress >= 30) || title.Contains("(review)"))
                {
                    item.BackColor = Color.DarkOrange;
                    item.ForeColor = Color.White;
                }
                if ((wordGoal != -1 && progress > 90) || title.Contains("(done)"))
                {
                    item.BackColor = Color.DarkGreen;
                    item.ForeColor = Color.White;
                }

                int fontSize = 10;
                switch (level)
                {
                    case 1:
                        fontSize = 12;
                        break;
                    case 2:
                        fontSize = 10;
                        break;
                    case 3:
                        fontSize = 9;
                        break;
                    default:
                        fontSize = 8;
                        break;
                }

                item.Font = new System.Drawing.Font("Arial", fontSize, GraphicsUnit.Point);

                listView1.Items.Add(item);
            }
        }

        private static int getLevel(Paragraph p)
        {
            int level = -1;

            try
            {
                level = Convert.ToInt32(p.OutlineLevel.ToString().Replace("wdOutlineLevel", ""));
            }
            catch (Exception) { }

            return level;
        }
    }
}
