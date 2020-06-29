using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.Runtime;
using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace Util
{
    /// <summary>
    /// 工具类
    /// </summary>
    public static class MyTools
    {
        public static void setDataGridView(this DataGridView dataGridView1,int rowNum,int colNum,double rowHeight,double colWidth,List<string> headString)
        {
            MessageBox.Show("设定表格!");
        }
        /// <summary>
        /// 限制输入整数
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        /// <returns></returns>
        public static bool inputint(object sender, KeyPressEventArgs e)
        {
            bool b = false;
            int keyCode = (int)e.KeyChar;
            if ((keyCode < 48 || keyCode > 57) && keyCode != 8 && keyCode != 13)
            {
                b = true;
            }
            return b;
        }
        /// <summary>
        /// 限制输入浮点数
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        /// <returns></returns>
        public static bool inputdouble(object sender, KeyPressEventArgs e)
        {
            bool b = false;
            int keyCode = (int)e.KeyChar;
            if ((keyCode < 48 || keyCode > 57) && keyCode != 8)
            {
                if (keyCode == 46)
                {
                    TextBox textbox = (TextBox)sender;
                    if (textbox.Text.IndexOf(".") != -1)
                        b = true;
                }
                else
                    b = true;
            }
            return b;
        }

        public static int Max(int a, int b)
        {
            if (a > b)
                return a;
            else
                return b;
        }
        //选择对象并返回DBObjectCollection
        public static DBObjectCollection HJ_GetObjsBySelectionFilter(this Database db, SelectionFilter filter, OpenMode mode, bool openErased)
        {
            Editor ed = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor;
            //选择符合条件的所有实体
            PromptSelectionResult entSelected = ed.GetSelection(filter);
            if (entSelected.Status != PromptStatus.OK) return null;
            SelectionSet ss = entSelected.Value;

            DBObjectCollection ents = new DBObjectCollection();

            using (Transaction ts = db.TransactionManager.StartTransaction())
            {
                foreach (ObjectId id in ss.GetObjectIds())
                {
                    DBObject obj = ts.GetObject(id, mode, openErased);
                    if (obj != null)
                        ents.Add(obj);
                }
                ts.Commit();//没有这个就不会提交内部的修改
            }
            return ents;
        }
        /// <summary>
        /// 重新定义Table行数列数
        /// </summary>
        /// <param name="table"></param>
        /// <param name="biaoTouRowsCount"></param>
        /// <param name="rowNum"></param>
        /// <param name="colNum"></param>
        /// <param name="rowHeight"></param>
        /// <param name="colWidth"></param>
        public static void ReSize(this Table table, int biaoTouRowsCount, int rowNum, int colNum, double rowHeight, double colWidth)
        {
            //列，插入、删除
            if (colNum > table.NumColumns)
            {
                if (table.CanInsertColumn(table.NumColumns))
                {
                    int ksCol = table.NumColumns;
                    table.InsertColumns(table.NumColumns, colWidth, colNum - table.NumColumns);
                    int jsCol = table.NumColumns;
                    if (table.NumRows >= 2)
                    {
                        for (int i = ksCol; i < jsCol; i++)
                        {
                            table.SetTextString(1, i, "列w_" + i);
                        }
                    }
                }
            }
            if (colNum < table.NumColumns)
            {
                if (table.CanDeleteColumns(table.NumColumns - (table.NumColumns - colNum), table.NumColumns - colNum))
                {
                    table.DeleteColumns(table.NumColumns - (table.NumColumns - colNum), table.NumColumns - colNum);
                }
            }
            //行，插入、删除
            if (rowNum > table.NumRows - biaoTouRowsCount)
            {
                if (table.CanInsertRow(table.NumRows))
                {
                    table.InsertRows(table.NumRows, rowHeight, rowNum - (table.NumRows - biaoTouRowsCount));
                }
            }
            if (rowNum < table.NumRows - biaoTouRowsCount)
            {
                if (table.CanDeleteRows(table.NumRows - (table.NumRows - rowNum - biaoTouRowsCount), table.NumRows - rowNum - biaoTouRowsCount))
                {
                    table.DeleteRows(table.NumRows - (table.NumRows - rowNum - biaoTouRowsCount), table.NumRows - rowNum - biaoTouRowsCount);
                }
            }
        }
        /// <summary>
        /// 重新定义DataGridView行数列数
        /// </summary>
        /// <param name="dataGridView"></param>
        /// <param name="biaoTouRowsCount"></param>
        /// <param name="rowNum"></param>
        /// <param name="colNum"></param>
        /// <param name="rowHeight"></param>
        /// <param name="colWidth"></param>
        public static void ReSize(this DataGridView dataGridView, int rowNum, int colNum)
        {
            //列，插入、删除
            while (colNum > dataGridView.Columns.Count)
            {
                int curColNum = dataGridView.Columns.Count;
                dataGridView.Columns.Add("列_" + curColNum.ToString(), "列_" + curColNum.ToString());
                //dataGridView.Columns[dataGridView.Columns.Count - 1].Width = colWidth;//默认列宽
            }
            while (colNum < dataGridView.Columns.Count)
            {
                int curColNum = dataGridView.Columns.Count;
                dataGridView.Columns.RemoveAt(curColNum - 1);
            }
            //行，插入、删除
            if (rowNum > dataGridView.Rows.Count)
            {
                int curRowNum = dataGridView.Rows.Count;//原行数
                int add = rowNum - curRowNum;

                string[] sValue = new string[dataGridView.Columns.Count];
                //为最后一行赋空值
                for (int i = 0; i < dataGridView.Columns.Count; i++)
                {
                    object obj = dataGridView[i, dataGridView.Rows.Count - 1].Value;
                    if (obj != null)
                        sValue[i] = obj.ToString();
                    else
                        sValue[i] = "";
                    dataGridView[i, dataGridView.Rows.Count - 1].Value = null;
                }
                for (int a = 0; a < add; a++)
                {
                    if (a == 0)
                        dataGridView.Rows.Add(sValue);
                    else
                        dataGridView.Rows.Add();
                }
            }
            if (rowNum < dataGridView.Rows.Count)
            {
                int curRowNum = dataGridView.Rows.Count;
                int cha = curRowNum - rowNum;
                //为最后一行重新赋值
                for (int i = 0; i < dataGridView.Columns.Count; i++)
                {
                    dataGridView[i, dataGridView.Rows.Count - 1].Value = dataGridView[i,curRowNum-cha-1].Value;
                }
                //从倒数第二行，倒着删除3行
                for (int i = 0; i < cha; i++)
                {
                    dataGridView.Rows.RemoveAt(curRowNum - cha-1);
                }
            }
            //禁止列头排序
            for (int i = 0; i < dataGridView.Columns.Count; i++)
            {
                dataGridView.Columns[i].SortMode = DataGridViewColumnSortMode.NotSortable;
            }
            dataGridView.Refresh();
        }

    }
}
