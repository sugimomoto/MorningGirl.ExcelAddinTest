using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Excel;
using System.Windows.Forms;
using Microsoft.Office.Tools.Ribbon;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Tooling.Connector;
using System.Configuration;

namespace MorningGirl.ExcelAddinTest
{
    public partial class Ribbon1
    {
        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {
            
        }
        
        /// <summary>
        /// CrmConnectionボタンのクリックイベント
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button1_Click(object sender, RibbonControlEventArgs e)
        {

            MessageBox.Show(string.Format("「{0}」に一致する取引先企業を検索します。", !string.IsNullOrEmpty(this.editBox_searchText.Text) ? this.editBox_searchText.Text : ""));
            
            // CrmServiceを生成
            var crmSvc = new CrmServiceClient(ConfigurationManager.ConnectionStrings["CRMOnline"].ConnectionString);

            // Dynamics CRM から取引先企業を取得するためのQueryExpressionを作成
            var query = new QueryExpression();
            query.EntityName = "account";
            query.ColumnSet.AllColumns = true;

            query.Criteria = new FilterExpression();
            query.Criteria.AddCondition("name", ConditionOperator.Like,
                !string.IsNullOrEmpty(this.editBox_searchText.Text) ? "%" + this.editBox_searchText.Text + "%" : "");
            
            var accounts = (EntityCollection)crmSvc.RetrieveMultiple(query);

            // GlobalsからThisAddinのインスタンスを辿って、ActiveSheetを取得
            var activeSheet = ((Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet);

            // 一つづつ取り出して、ExcelのA列に格納
            foreach (var account in accounts.Entities.Select((item,i) => new { item, i }))
            {
                var resultCell = (Excel.Range)activeSheet.get_Range("A" + (account.i + 1));
                resultCell.Value2 = account.item["name"];
            }
        }
    }
}
