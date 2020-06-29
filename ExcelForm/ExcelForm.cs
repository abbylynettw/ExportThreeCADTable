using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.Runtime;
using System.Runtime.InteropServices;
using acadApp = Autodesk.AutoCAD.ApplicationServices.Application;
using winApp = System.Windows.Forms;
using DotNetARX;
using Util;
using MyEntity;

namespace ExcelForm
{
    public partial class ExcelForm : Form
    {
        //判断鼠标是否位于窗体/控件内
        [DllImport("user32.dll", EntryPoint = "GetWindowRect")]
        static extern int GetWindowRect(IntPtr hwnd, out Rect lpRect);
        public struct Rect
        {
            public int Left;
            public int Top;
            public int Right;
            public int Bottom;
        }
        [DllImport("user32.dll", EntryPoint = "GetCursorPos")]
        static extern bool GetCursorPos(ref Point lpPoint);
        [DllImport("user32.dll", EntryPoint = "PtInRect")]
        static extern int PtInRect(Rect lpRect, Point pt);
        //激活CAD用
        [DllImport("user32.dll", EntryPoint = "SetFocus")]
        public static extern int SetFocus(IntPtr hWnd);

        ObjectId tableId = ObjectId.Null;
        int tableCount = 0;//插入的表格数量
        int biaoTouRowsCount = 2;
        int columnHeaderIndex;
        bool runCmd = false;
        private int dataRowCount;
        bool autoChange = false;
        Database mDatabase = HostApplicationServices.WorkingDatabase;

        public ExcelForm()
        {
            InitializeComponent();
        }
        private void ExcelForm_Load(object sender, EventArgs e)
        {
            /*
            //初始化ExcelClass 好像多余
            ExcelClass.rowNum = 0;
            ExcelClass.colNum = 0;
            ExcelClass.selectRow = 0;
            ExcelClass.selectCol = 0;
            ExcelClass.headString = null;
            */

            //增加对象修改事件
            mDatabase.ObjectModified += objectModified;

            //文本框的值已在窗体设计时加入了默认值
            setExcelClassValue();
            dataGridView1.ReSize(ExcelClass.rowNum, ExcelClass.colNum);//初始化

            //第一次运行之后CAD会记住窗体的位置和大小，然后下面两行代码就不起作用了。
            this.Width = 817;
            this.Height = 500;
            ExcelForm_Resize(sender, e);
        }

        private void objectModified(object sender, ObjectEventArgs e)
        {
            //如果表格ID != ObjectId.Null
            if (tableId != ObjectId.Null)
            {
                //如果修改的对象ID==tableId（当前表格ID）
                if (e.DBObject.Id == tableId)
                {
                    //避免递归，取消事件★★★
                    mDatabase.ObjectModified -= objectModified;
                    try
                    {
                        if (autoChange == false)
                        {
                            this.dataGridView1_Fill(tableId);
                        }
                    }
                    catch (System.Exception ex)
                    {
                        MessageBox.Show("发生错误：autoChange变量值可能不是false!");
                    }
                    finally
                    {
                        //重新启动事件
                        mDatabase.ObjectModified += objectModified;
                    }
                }
            }
        }

        private void ExcelForm_Resize(object sender, EventArgs e)
        {
            dataGridView1.Left = 10;
            dataGridView1.Top= 50;
            if (this.Height > 200)
                dataGridView1.Height = this.Height - 100;
            else
                dataGridView1.Height = 100;
            if (this.Width > 200)
                dataGridView1.Width = this.Width - 25;
            else
                dataGridView1.Width = 175;
        }
        /// <summary>
        /// 获取文字列表
        /// </summary>
        /// <param name="objlist"></param>
        /// <returns></returns>
        private List<string> get_text_list(DBObjectCollection objlist)
        {
            List<string> list = new List<string>();
            foreach (DBObject obj in objlist)
            {
                string dxfName = RXClass.GetClass(obj.GetType()).DxfName;
                string txt = "";
                switch (dxfName)
                {
                    case "TEXT":
                        DBText txtobj = obj as DBText;
                        txt = txtobj.TextString;
                        list.Add(txt);
                        break;
                    case "MTEXT":
                        MText mtxt = obj as MText;
                        string txtall = mtxt.Contents;
                        string[] sarr = txtall.Split(new string[] { "\\P" }, StringSplitOptions.None);

                        if (checkBox_dhcl.Checked == true)
                        {
                            sarr.ToList().ForEach(a => list.Add(a));
                        }
                        else
                        {
                            list.Add(List_strcat(sarr.ToList()));
                        }
                        break;
                    case "DIMENSION":
                        Entity ent = obj as Entity;
                        DBObjectCollection objs1 = new DBObjectCollection();
                        ent.Explode(objs1);
                        List<string> rsStr1 = get_text_list(objs1);
                        list.AddRange(rsStr1);
                        break;
                    case "INSERT":
                        BlockReference blk = obj as BlockReference;
                        DBObjectCollection objs = new DBObjectCollection();
                        blk.Explode(objs);
                        List<string> rsStr = get_text_list(objs);
                        list.AddRange(rsStr);
                        break;
                    default:
                        break;
                }
            }
            return list;
        }
        /// <summary>
        /// 列表内文字合并为一个字符串
        /// </summary>
        /// <param name="rsStr"></param>
        /// <returns></returns>
        private string List_strcat(List<string> rsStr)
        {
            string rss = "";
            foreach (string s in rsStr)
            {
                rss += (s + " ");
            }
            return rss.TrimEnd();
        }
        /// <summary>
        /// 确定按钮，插入表格
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button_define_Click_1(object sender, EventArgs e)
        {
            if (runCmd == true) return;//命令不重复执行
            runCmd = true;

            //有必要再检查一遍，比如行列从5改为1，点否，然后没有恢复>=5。
            bool bl = checkTextbox();
            if (!bl)
            {
                runCmd = false;
                return;//数据填写有误时，退出。
            }

            autoChange = true;//自动修改表格

            //如果确定缩小表格，重定义ExcelClass
            setExcelClassValue();//设置ExcelClass类的属性
            
            //激活CAD窗口
            SetFocus(acadApp.DocumentManager.MdiActiveDocument.Window.Handle);

            Document doc = acadApp.DocumentManager.MdiActiveDocument;
            Editor ed = doc.Editor;
            Database db = doc.Database;
            
            //创建表格
            PromptPointOptions ppo = new PromptPointOptions("\n指定表格插入点:");
            ppo.AllowNone = true;
            PromptPointResult ppr = ed.GetPoint(ppo);
            if (ppr.Status != PromptStatus.OK) return;
            Point3d pt = ppr.Value;
            tableId = createTable(db, "标题行",pt,true);

            //更新DataGridView
            dataGridView1_Fill(tableId);

            //设置当前行 当前列
            ExcelClass.selectCol = dataGridView1.CurrentCell.ColumnIndex;
            ExcelClass.selectRow = dataGridView1.CurrentCell.RowIndex;

            tableCount++;//插入的表格数量
            runCmd = false;

            autoChange = false;
        }
        //填充DataGridView内容
        private void dataGridView1_Fill(ObjectId tableId)
        {
            Database db = HostApplicationServices.WorkingDatabase;
            using (Transaction trans = db.TransactionManager.StartTransaction())
            {
                Table table = trans.GetObject(tableId, OpenMode.ForRead, false) as Table;
                //一步设定dataGridView1大小=table大小
                dataGridView1.ReSize(table.NumRows - biaoTouRowsCount, table.NumColumns);
                //设置列 表头
                for (int i = 0; i < table.NumColumns; i++)
                {
                    string bt = table.Value(1, i) as string;
                    dataGridView1.Columns[i].Name = bt;
                    dataGridView1.Columns[i].HeaderText = bt;
                }
                //设置行 数据
                for (int i = 0; i < table.NumRows-biaoTouRowsCount; i++)
                {
                    string[] sArr = new string[table.NumColumns];
                    for (int j = 0; j < table.NumColumns; j++)
                    {
                        string s1 = table.GetTextString(i + biaoTouRowsCount, j, 0);
                        sArr[j] = s1;
                    }
                    dataGridView1.Rows[i].SetValues(sArr);
                }
                dataGridView1.Refresh();

                //重新设置行数(如果发生了变化)，以下4行主要作用于autoChang==false时
                ExcelClass.rowNum = table.NumRows - biaoTouRowsCount;
                ExcelClass.colNum = table.NumColumns;
                textBox_row.Text = ExcelClass.rowNum.ToString();
                textBox_col.Text = ExcelClass.colNum.ToString();

                trans.Commit();
            }
        }

        /// <summary>
        /// 检查行数，列数，是否为大于0的整数，检查行高，列宽是否为大于0的浮点。
        /// </summary>
        /// <returns></returns>
        private bool checkTextbox()
        {
            bool bl;
            int i;
            double b;
            //行数
            bl= int.TryParse(textBox_row.Text , out i);
            if (bl)
            {
                if (i > 0)
                {
                    if (i < ExcelClass.rowNum)
                    {
                        DialogResult dr = MessageBox.Show("设置的行数("+ i.ToString() + ")小于原有行数(" + ExcelClass.rowNum.ToString() + ")，继续运行可能会删除数据，是否继续？", "提示", MessageBoxButtons.YesNo);
                        if (dr == DialogResult.No)
                            return false;
                    }
                }
                else
                {
                    MessageBox.Show("行数应大于0！");
                    return false;
                }
            }
            else
            {
                MessageBox.Show("行数应为数值！");
                return false;
            }
            //列数
            bl = int.TryParse(textBox_col.Text, out i);
            if (bl)
            {
                if (i > 0)
                {
                    if (i < ExcelClass.colNum)
                    {
                        DialogResult dr = MessageBox.Show("设置的列数(" + i.ToString() + ")小于原有列数(" + ExcelClass.colNum.ToString() + ")，继续运行可能会删除数据，是否继续？", "提示", MessageBoxButtons.YesNo);
                        if (dr == DialogResult.No)
                            return false;
                    }
                }
                else
                {
                    MessageBox.Show("列数应大于0！");
                    return false;
                }
            }
            else
            {
                MessageBox.Show("列数应为数值！");
                return false;
            }
            
            //行高
            bl = double.TryParse(textBox_rowHeight.Text, out b);
            if (bl)
            {
                if (b <= 0)
                {
                    MessageBox.Show("行高应大于0！");
                    return false;
                }
            }
            else
            {
                MessageBox.Show("行高应为数值！");
                return false;
            }
            //列宽
            bl = double.TryParse(textBox_colWidth.Text , out b);
            if (bl)
            {
                if (b <= 0)
                {
                    MessageBox.Show("列宽应大于0！");
                    return false;
                }
            }
            else
            {
                MessageBox.Show("列宽应为数值！");
                return false;
            }
            return true;
        }
        //根据TextBox设置ExcelClass
        private void setExcelClassValue()
        {
            ExcelClass.rowNum = int.Parse(textBox_row.Text);
            ExcelClass.colNum = int.Parse(textBox_col.Text);
            ExcelClass.rowHeight = double.Parse(textBox_rowHeight.Text);
            ExcelClass.colWidth = double.Parse(textBox_colWidth.Text);
            if (ExcelClass.headString == null)
                ExcelClass.headString = new List<string>();
            else
                ExcelClass.headString.Clear();

            if (tableCount==0)
            {
                //初始化时运行一次,直接运行else应该也行
                for (int i = 0; i < ExcelClass.colNum; i++)
                {
                    ExcelClass.headString.Add("列_" + i.ToString());
                }
                tableCount++;
            }
            else
            {
                int colCount = dataGridView1.Columns.Count;
                string[] btArr = new string[colCount];
                for (int i = 0; i < colCount; i++)
                {
                    btArr[i] = dataGridView1.Columns[i].Name;
                }
                //修改表头List
                for (int i = 0; i < ExcelClass.colNum; i++)
                {
                    if(i<btArr.Length)
                        ExcelClass.headString.Add(btArr[i]);
                    else
                        ExcelClass.headString.Add("列_" + i.ToString());
                }
            }
        }
        /// <summary>
        /// 创建表格
        /// </summary>
        /// <param name="db"></param>
        /// <param name="title"></param>
        /// <param name="insertPt"></param>
        /// <param name="insToPt"></param>
        /// <returns></returns>
        private ObjectId createTable(Database db,string title,Point3d insertPt,bool insToPt)
        {
            ObjectId objID = ObjectId.Null;
            using (DocumentLock doclock = acadApp.DocumentManager.MdiActiveDocument.LockDocument())
            {
                using (Transaction trans = db.TransactionManager.StartTransaction())
                {
                    ObjectId styleId = AddTableStyle("MyTable");//get样式，若未定义则新建
                    Table table = new Table();
                    table.TableStyle = styleId;
                    if (insToPt == true)
                        table.Position = insertPt;
                    setTableRowColHeader(table, title);//设置行数，列数，行高，列宽，表头
                    objID = db.AddToModelSpace(table);
                    trans.Commit();
                }
            }
            return objID;
        }
        /// <summary>
        /// 设定Table行数，列数，行高，列宽，表头
        /// </summary>
        /// <param name="table"></param>
        /// <param name="title"></param>
        private void setTableRowColHeader(Table table, string title)
        {
            table.ReSize(biaoTouRowsCount, ExcelClass.rowNum, ExcelClass.colNum, ExcelClass.rowHeight, ExcelClass.colWidth);
            table.SetRowHeight(ExcelClass.rowHeight);     //设定行高
            table.SetColumnWidth(ExcelClass.colWidth);    //设定列宽
            table.SetTextHeight(ExcelClass.rowHeight * 0.8, TableTools.AllRows);   //重新设定文字高度=rowHeight*0.8
            if(title != "")
                table.SetTextString(0, 0, title);//初始化后就不再定义
            List<string> btList = ExcelClass.headString as List<string>;
            for (int i = 0; i < btList.Count; i++)
            {
                //设置表头文字
                table.SetTextString(1,i, btList[i]);
            }
        }
        /// <summary>
        /// 为当前图形添加一个新的表格样式
        /// </summary>
        /// <param name="style"></param>
        /// <returns></returns>
        public static ObjectId AddTableStyle(string style)
        {
            ObjectId styleId; // 存储表格样式的Id
            Document doc = acadApp.DocumentManager.MdiActiveDocument;
            Database db = HostApplicationServices.WorkingDatabase;
            using (DocumentLock doclock = doc.LockDocument())
            {
                using (Transaction trans = db.TransactionManager.StartTransaction())
                {
                    // 打开表格样式字典
                    DBDictionary dict = (DBDictionary)db.TableStyleDictionaryId.GetObject(OpenMode.ForRead);
                    if (dict.Contains(style)) // 如果存在指定的表格样式
                    {
                        styleId = dict.GetAt(style); // 获取表格样式的Id
                    }
                    else
                    {
                        TableStyle ts = new TableStyle(); // 新建一个表格样式
                        // 设置表格所有行的外边框的线宽为0.30mm
                        ts.SetGridLineWeight(LineWeight.LineWeight030, (int)GridLineType.OuterGridLines, TableTools.AllRows);
                        // 不加粗表格表头行的底部边框
                        ts.SetGridLineWeight(LineWeight.LineWeight000, (int)GridLineType.HorizontalBottom, (int)RowType.HeaderRow);
                        // 不加粗表格数据行的顶部边框
                        ts.SetGridLineWeight(LineWeight.LineWeight000, (int)GridLineType.HorizontalTop, (int)RowType.DataRow);
                        // 设置表格中所有行的文本高度为1(默认)
                        ts.SetTextHeight(1, TableTools.AllRows);
                        // 设置表格中所有行的对齐方式为正中
                        ts.SetAlignment(CellAlignment.MiddleCenter, TableTools.AllRows);
                        dict.UpgradeOpen();//切换表格样式字典为写的状态
                        
                        // 将新的表格样式添加到样式字典并获取其Id
                        styleId = dict.SetAt(style, ts);
                        // 将新建的表格样式添加到事务处理中
                        trans.AddNewlyCreatedDBObject(ts, true);
                        dict.DowngradeOpen();
                        trans.Commit();
                    }
                }
            }
            return styleId; // 返回表格样式的Id
        }
        /// <summary>
        /// 修改Excelclass当前行，当前列
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void dataGridView1_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            ExcelClass.selectCol = dataGridView1.CurrentCell.ColumnIndex;
            ExcelClass.selectRow = dataGridView1.CurrentCell.RowIndex;
        }
        /// <summary>
        /// 选择文字（按钮）
        /// </summary>
        private void button_selt_Click(object sender, EventArgs e)
        {
            //如果表格不存在，提示绘制表格。
            if (tableId == ObjectId.Null)
            {
                acadApp.ShowAlertDialog("请先绘制表格!");
                return;
            }

            //由↑↓←→键造成的变量不跟随改变问题。下边两行解决
            ExcelClass.selectCol = dataGridView1.CurrentCell.ColumnIndex;
            ExcelClass.selectRow = dataGridView1.CurrentCell.RowIndex;

            autoChange = true;//自动修改表格

            Document doc = acadApp.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            Editor ed = doc.Editor;
            
            //判断是按行填充，还是按列填充。
            using (DocumentLock doclock = doc.LockDocument())
            {
                using (Transaction trans = db.TransactionManager.StartTransaction())
                {
                    try
                    {
                        Table table = trans.GetObject(tableId, OpenMode.ForWrite, false) as Table;
                        //激活CAD文档
                        SetFocus(acadApp.DocumentManager.MdiActiveDocument.Window.Handle);
                        //构建选择集过滤器   
                        TypedValue[] values = { new TypedValue((int)DxfCode.Start, "*TEXT,INSERT,DIMENSION"),
                                    new TypedValue((int)DxfCode.LayoutName,"Model")
                                  };
                        SelectionFilter filter = new SelectionFilter(values);

                        DBObjectCollection objlist = db.HJ_GetObjsBySelectionFilter(filter, OpenMode.ForRead, false);
                        if (objlist == null) return;//未选择对象
                        List<string> list = get_text_list(objlist);
                        int txtCount = list.Count;
                        
                        //判断填充方式
                        if (radio_row.Checked == true)
                        {
                            //按行填充
                            
                            //循环赋值
                            foreach (string str in list)
                            {
                                //str in list
                                //判断单元格是否为空，不为空是否覆盖。
                                if (table.GetTextString(ExcelClass.selectRow + biaoTouRowsCount, ExcelClass.selectCol, 0) != "")
                                {
                                    DialogResult dr = MessageBox.Show("单元格不为空，是否覆盖？", "提示!", MessageBoxButtons.YesNo);
                                    if (dr != DialogResult.No)
                                    {
                                        table.SetTextString(ExcelClass.selectRow + biaoTouRowsCount, ExcelClass.selectCol, str);
                                        dataGridView1[ExcelClass.selectCol, ExcelClass.selectRow].Value = str;
                                    }
                                }
                                else
                                {
                                    table.SetTextString(ExcelClass.selectRow + biaoTouRowsCount, ExcelClass.selectCol, str);
                                    dataGridView1[ExcelClass.selectCol, ExcelClass.selectRow].Value = str;
                                }
                                //当前列+1
                                ExcelClass.selectCol += 1;
                                txtCount--;//数量-1

                                if (ExcelClass.selectCol == dataGridView1.Columns.Count)
                                {
                                    ExcelClass.selectCol = 0;
                                    ExcelClass.selectRow += 1;
                                    //if达到末尾行并且还有文字
                                    if (ExcelClass.selectRow  ==  dataGridView1.Rows.Count && txtCount>0)
                                    {
                                        //table新建行
                                        ExcelClass.rowNum += 1;
                                        table.ReSize(biaoTouRowsCount, ExcelClass.rowNum, ExcelClass.colNum, ExcelClass.rowHeight, ExcelClass.colWidth);
                                        //form新建行  需要修改
                                        dataGridView1.ReSize(ExcelClass.rowNum, ExcelClass.colNum);
                                    }
                                }
                                //因为FORM表格多一行，所以[ExcelClass.selectCol, ExcelClass.selectRow]这个位置不会出错
                                //按列填充时，这个位置selectCol要-1才行，不然会产生异常。
                                if (ExcelClass.selectRow == dataGridView1.Rows.Count && txtCount == 0)
                                {
                                    ExcelClass.selectRow--;
                                    ExcelClass.selectCol = dataGridView1.Columns.Count - 1;
                                    //winform填到最后一格了（没有新增行）
                                }
                                dataGridView1.CurrentCell = dataGridView1[ExcelClass.selectCol, ExcelClass.selectRow];
                                dataGridView1.Refresh();
                            }
                        }
                        else
                        {
                            //按列填充

                            //循环赋值
                            foreach (string str in list)
                            {
                                //str in list
                                //判断单元格是否为空，不为空是否覆盖
                                if (table.GetTextString(ExcelClass.selectRow + biaoTouRowsCount, ExcelClass.selectCol, 0) != "")
                                {
                                    DialogResult dr = MessageBox.Show("单元格不为空，是否覆盖？", "提示!", MessageBoxButtons.YesNo);
                                    if (dr != DialogResult.No)
                                    {
                                        table.SetTextString(ExcelClass.selectRow + biaoTouRowsCount, ExcelClass.selectCol, str);
                                        dataGridView1[ExcelClass.selectCol, ExcelClass.selectRow].Value = str;
                                    }
                                }
                                else
                                {
                                    table.SetTextString(ExcelClass.selectRow + biaoTouRowsCount, ExcelClass.selectCol, str);//可能要增加行
                                    dataGridView1[ExcelClass.selectCol, ExcelClass.selectRow].Value = str;
                                }
                                
                                ExcelClass.selectRow += 1;
                                txtCount--;//数量-1

                                if (ExcelClass.selectRow  == dataGridView1.Rows.Count)
                                {
                                    ExcelClass.selectRow  = 0;
                                    ExcelClass.selectCol += 1;
                                    //if达到末尾行并且还有文字
                                    if (ExcelClass.selectCol == dataGridView1.Columns.Count && txtCount > 0)
                                    {
                                        //table新建列
                                        ExcelClass.colNum += 1;
                                        table.ReSize(biaoTouRowsCount, ExcelClass.rowNum, ExcelClass.colNum, ExcelClass.rowHeight, ExcelClass.colWidth);
                                        //form 新建列
                                        dataGridView1.Columns.Add("列_" + (ExcelClass.colNum-1).ToString(), "列_" + (ExcelClass.colNum - 1).ToString());
                                        ExcelClass.headString.Add("列_" + (ExcelClass.colNum - 1).ToString());
                                    }
                                }
                                //因为FORM表格多一行，所以[ExcelClass.selectCol, ExcelClass.selectRow]这个位置不会出错
                                //按列填充时，这个位置selectCol要-1才行，不然会产生异常。
                                if (ExcelClass.selectCol == dataGridView1.Columns.Count && txtCount == 0)
                                {
                                    ExcelClass.selectCol--;
                                    ExcelClass.selectRow = dataGridView1.Rows.Count-1;
                                    //winform填到最后一格了（没有新增列）
                                }
                                dataGridView1.CurrentCell = dataGridView1[ExcelClass.selectCol, ExcelClass.selectRow];
                                dataGridView1.Refresh();
                            }
                        }//if 填充方式 == 按行，按列。
                        dataGridView1.Refresh();
                        table.DowngradeOpen();
                        trans.Commit();
                    }
                    catch (Autodesk.AutoCAD.Runtime.Exception ex)
                    {
                        trans.Abort();
                        if (ex.Message == "eWasErased")
                        {
                            acadApp.ShowAlertDialog("对象已被删除，请重新插入表格!");
                            tableId = ObjectId.Null;
                        }
                        else
                        {
                            acadApp.ShowAlertDialog("错误：" + ex.Message);
                        }
                    }
                }//trans
            }//doc.lock

            //填充之后表格可能扩大了，文本框内容根据ExcelClass的行数列数改变。
            textBox_row.Text = ExcelClass.rowNum.ToString();
            textBox_col.Text = ExcelClass.colNum.ToString();

            autoChange = false;
        }
        private void textBox_row_Leave(object sender, EventArgs e)
        {
            thisApp_update();
        }
        private void textBox_row_KeyPress(object sender, KeyPressEventArgs e)
        {
            //限制输入整型数
            e.Handled = MyTools.inputint(sender, e);
            if ((int)e.KeyChar == 13)
                thisApp_update();
        }
        private void textBox_col_Leave(object sender, EventArgs e)
        {
            thisApp_update();
        }
        private void textBox_col_KeyPress(object sender, KeyPressEventArgs e)
        {
            //限制输入整型数
            e.Handled = MyTools.inputint(sender, e);
            if ((int)e.KeyChar == 13)
                thisApp_update();
        }

        private void textBox_rowHeight_TextChanged(object sender, EventArgs e)
        {
            thisApp_update();
        }
        private void textBox_rowHeight_KeyPress(object sender, KeyPressEventArgs e)
        {
            //限制输入浮点数
            e.Handled = MyTools.inputdouble(sender, e);
        }
        private void textBox_colWidth_TextChanged(object sender, EventArgs e)
        {
            thisApp_update();
        }
        /// <summary>
        /// 更新dataGridView控件
        /// </summary>
        private void thisApp_update()
        {
            autoChange = true;//自动修改表格

            //输入类型已加入限制，无需再次判断输入类型，需要判断是否变小了
            bool bl = checkTextbox();
            if (!bl) return;//数据填写有误时，退出。
            setExcelClassValue();
            dataGridView1.ReSize(ExcelClass.rowNum, ExcelClass.colNum);
            tableLianDong("外观");
            ExcelClass.selectCol = dataGridView1.CurrentCell.ColumnIndex;
            ExcelClass.selectRow = dataGridView1.CurrentCell.RowIndex;

            autoChange = false;
        }
        private void textBox_colWidth_KeyPress(object sender, KeyPressEventArgs e)
        {
            //限制输入浮点数
            e.Handled = MyTools.inputdouble(sender, e);
        }
        /// <summary>
        /// dataGridView更新后，table也更新(leixing=="外观"时，设置表格大小，leixing=="数据"时，设定数据内容。)
        /// </summary>
        private void tableLianDong(string leixing)
        {
            tableLianDong(leixing, 0, 0, "");
        }
        private void tableLianDong(string leixing,int row, int col, string value)
        {
            if (tableCount > 0)
            {
                Document doc = acadApp.DocumentManager.MdiActiveDocument;
                Database db = doc.Database;
                //修改表格
                using (DocumentLock doclock = doc.LockDocument())
                {
                    using (Transaction trans = db.TransactionManager.StartTransaction())
                    {
                        try
                        {
                            Table table = trans.GetObject(tableId, OpenMode.ForWrite) as Table;
                            if(leixing == "外观")
                                setTableRowColHeader(table, "");//设置Table行数，列数，行高，列宽，(表头不变)
                            if (leixing == "数据")
                                table.SetTextString(row + biaoTouRowsCount, col, value);
                            //本身联动应该设置value值，但能调用此函数的函数，无需设置value值。
                            //设置Value值的功能，应定义到dataGridView1_cellChange事件里边。
                            table.DowngradeOpen();
                            trans.Commit();
                        }
                        catch (Autodesk.AutoCAD.Runtime.Exception ex)
                        {
                            trans.Abort();
                            if (ex.Message == "eWasErased")
                            {
                                acadApp.ShowAlertDialog("对象已被删除，请重新插入表格!");
                                tableId = ObjectId.Null;
                            }
                        }
                    }
                }
            }
        }
        /// <summary>
        /// 双击DataGridView表头
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void dataGridView1_ColumnHeaderMouseDoubleClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            bool bl = checkTextbox();
            if (!bl) return;//数据填写有误时，退出。

            columnHeaderIndex = e.ColumnIndex;
            string biaotou = dataGridView1.Columns[columnHeaderIndex].Name;
            textBox_HeaderNameEdit.Location = this.PointToClient(Cursor.Position);
            textBox_HeaderNameEdit.Text = biaotou;
            textBox_HeaderNameEdit.Visible = true;
            textBox_HeaderNameEdit.Focus();//使其获得焦点
        }
        /// <summary>
        /// 编辑表头
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void textBox_HeaderNameEdit_KeyPress(object sender, KeyPressEventArgs e)
        {
            if ((int)e.KeyChar == 13)
            {
                autoChange = true;//自动修改
                //回车
                if (textBox_HeaderNameEdit.Text != "")
                {
                    dataGridView1.Columns[columnHeaderIndex].Name = textBox_HeaderNameEdit.Text.Trim();
                    dataGridView1.Columns[columnHeaderIndex].HeaderText = textBox_HeaderNameEdit.Text.Trim();
                    textBox_HeaderNameEdit.Visible = false;
                    setExcelClassValue();
                    //CAD表格联动
                    tableLianDong("外观");
                }
                autoChange = false;
            }
            if((int)e.KeyChar == 27)
            {
                //取消
                textBox_HeaderNameEdit.Visible = false;
            }
        }

        private void textBox_HeaderNameEdit_Leave(object sender, EventArgs e)
        {
            textBox_HeaderNameEdit.Visible = false;
        }
        private void dataGridView1_KeyPress(object sender, KeyPressEventArgs e)
        {
            //MessageBox.Show(dataGridView1.Focused.ToString());
            //通过上一行代码得出结论：如果当前单元格已是右下角最后一个，则本函数不会运行。
            //所以我们要到（textbox_row）的获得焦点事件中去判断，如果获得焦点的时候，tabRow，tabCol都=最大值 && 鼠标不在文本框位置，则新建行/列。
            //
            //判断按了哪个键，如果为enter键，取消功能(取消不了)。如果为DELETE键（DELETE获取不到），删除数据，并更新(整个表格,还用那个联动函数。)。
            //dataGridView1.SelectedCells[0].Value = "";

            //回车键默认系统功能（本函数只能接收回车和TAB键...）
            if (((int)e.KeyChar) == 13)
                return;

            //TAB键功能自定义
            try
            {
                if (radio_row.Checked == true)
                {
                    //按行,使用系统自带功能，否则会跳两个格。
                    /*
                    if (ExcelClass.selectCol < dataGridView1.Columns.Count - 1)
                    {
                        int curRow = dataGridView1.CurrentCell.RowIndex;
                        int curCol = dataGridView1.CurrentCell.ColumnIndex;
                        ExcelClass.selectCol = curCol + 1;
                        ExcelClass.selectRow = curRow;
                        dataGridView1.CurrentCell = dataGridView1[curCol + 1, curRow];//先列，后行
                    }
                    */
                }
                else
                {
                    //按列
                    if (ExcelClass.selectRow < dataGridView1.Rows.Count - 1)
                    {
                        int curRow = dataGridView1.CurrentCell.RowIndex;
                        int curCol = dataGridView1.CurrentCell.ColumnIndex;
                        if (curCol == 0)
                        {
                            curCol = dataGridView1.Columns.Count;
                            curRow--;
                        }
                        dataGridView1.CurrentCell = dataGridView1[curCol - 1, curRow + 1];//先列，后行
                    }
                    
                    if (ExcelClass.selectRow == dataGridView1.Rows.Count - 1)
                    {
                        int curRow = dataGridView1.CurrentCell.RowIndex;
                        int curCol = dataGridView1.CurrentCell.ColumnIndex;
                        
                        if (curCol < dataGridView1.Columns.Count)
                        {
                            curRow = 0;
                            curCol = curCol;
                        }
                        /*else
                        {
                            //结尾，焦点会转移到其他控件（不会进入此函数运行）
                            curRow = dataGridView1.Rows.Count-1;
                            curCol = dataGridView1.Columns.Count - 1;
                        }*/
                        dataGridView1.CurrentCell = dataGridView1[curCol, curRow];//先列，后行
                    }
                }
            }
            catch (System.Exception ex)
            {

            }
            //无论如何，selectCol，selectRow要根据显示焦点而定义!
            ExcelClass.selectCol = dataGridView1.CurrentCell.ColumnIndex;
            ExcelClass.selectRow = dataGridView1.CurrentCell.RowIndex;
        }
        private void dataGridView1_CellBeginEdit(object sender, DataGridViewCellCancelEventArgs e)
        {
            dataRowCount = dataGridView1.Rows.Count;
        }
        private void dataGridView1_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {
            int dataRowCount1  = dataGridView1.Rows.Count;

            if (dataRowCount1 != dataRowCount)
            {
                //MessageBox.Show("行数发生了变化! 将修改textBox_row");
                textBox_row.Text = dataRowCount1.ToString();//不会自动调用thisApp_update事件
                thisApp_update();
            }

            autoChange = true;//自动修改

            //更新Table表格的数据
            int row = dataGridView1.CurrentCell.RowIndex;
            int col = dataGridView1.CurrentCell.ColumnIndex;
            try
            {
                string str;
                object o = dataGridView1.CurrentCell.Value;
                if (o == null)
                    str = "";
                else
                    str = o.ToString();
                tableLianDong("数据", row, col, str);//数据联动，配备其他参数。
            }
            catch (System.Exception ex)
            {
                //MessageBox.Show("发生异常："+ ex.Message);
            }
            finally
            {
                autoChange = false;
            }
        }
        /// <summary>
        /// 文本框获得焦点时
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void textBox_row_Enter(object sender, EventArgs e)
        {
            autoChange = true;//自动修改表格

            //如果当前单元格为最后一个，并且TAB（行列）位置也是最后一个
            int curRow = dataGridView1.CurrentCell.RowIndex;
            int curCol = dataGridView1.CurrentCell.ColumnIndex;

            if (curRow == dataGridView1.Rows.Count - 1 && curCol == dataGridView1.Columns.Count - 1)
            {
                if (mouseInControl(this.textBox_row.Handle) == true)
                {
                    return;
                }
                else
                {
                    if (radio_row.Checked == true)
                    {
                        //按行
                        //MessageBox.Show("增加行并激活dataGridView");
                        //
                        dataGridView1.Focus();

                        ExcelClass.selectCol = 0;
                        ExcelClass.selectRow += 1;
                        ExcelClass.rowNum += 1;

                        //table新建行
                        dataGridView1.ReSize(ExcelClass.rowNum, ExcelClass.colNum);
                        dataGridView1.CurrentCell = dataGridView1[ExcelClass.selectCol, ExcelClass.selectRow];
                        tableLianDong("外观");
                        textBox_row.Text = "" + ExcelClass.rowNum;
                    }
                    else
                    {
                        //按列
                        //MessageBox.Show("增加列并激活dataGridView");
                        //
                        dataGridView1.Focus();

                        ExcelClass.selectRow = 0;
                        ExcelClass.selectCol += 1;
                        ExcelClass.colNum += 1;

                        //table新建列
                        dataGridView1.ReSize(ExcelClass.rowNum, ExcelClass.colNum);
                        dataGridView1.CurrentCell = dataGridView1[ExcelClass.selectCol, ExcelClass.selectRow];
                        tableLianDong("外观");
                        textBox_col.Text = "" + ExcelClass.colNum;
                    }
                }
            }
            autoChange = false;
        }
        /// <summary>
        /// 判断鼠标是否在TextBoxRow控件内部
        /// </summary>
        /// <returns></returns>
        private bool mouseInControl(IntPtr handle)
        {
            Point pt = new Point();
            bool b = GetCursorPos(ref pt);
            if (b)
            {
                Rect r = new Rect();
                int i = GetWindowRect(handle, out r);
                if (i != 0)
                {
                    if (PtInRect(r, pt) != 0)
                        return true;
                    else
                        return false;
                }
                else
                    return false;
                
            }
            else
                return false;
        }
        private void dataGridView1_CurrentCellChanged(object sender, EventArgs e)
        {
            //用这个事件的话，TAB特别定制功能将出错。
        }

        private void ExcelForm_FormClosed(object sender, FormClosedEventArgs e)
        {
            //删除对象修改事件
            mDatabase.ObjectModified -= objectModified;
        }

        private void ExcelForm_MouseLeave(object sender, EventArgs e)
        {
            //如果鼠标移出窗体范围，则视为手工修改表格。
            //if (mouseInControl(this.Handle) == false)
                
        }
    }
}
