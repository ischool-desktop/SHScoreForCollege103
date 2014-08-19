using FISCA;
using FISCA.Permission;
using FISCA.Presentation;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace SHScoreForCollege103
{
    public class Program
    {
        [MainMethod()]
        public static void Main()
        {
            // 匯入高關懷特殊身分
            Catalog catalog01 = RoleAclSource.Instance["學生"]["報表"];
            catalog01.Add(new RibbonFeature("SHScoreForCollege103.ScoreForm", "103學年度大學繁星推甄成績檔"));

            RibbonBarItem item01 = K12.Presentation.NLDPanels.Student.RibbonBarItems["資料統計"];
            item01["報表"]["成績相關報表"]["103學年度大學繁星推甄成績檔"].Enable = UserAcl.Current["SHScoreForCollege103.ScoreForm"].Executable;
            item01["報表"]["成績相關報表"]["103學年度大學繁星推甄成績檔"].Click += delegate
            {
                if (K12.Presentation.NLDPanels.Student.SelectedSource.Count > 0)
                {
                    ScoreForm sf = new ScoreForm();
                    sf.ShowDialog();
                }
                else
                {
                    FISCA.Presentation.Controls.MsgBox.Show("請選擇學生!");
                }
            };
        }
    }
}
