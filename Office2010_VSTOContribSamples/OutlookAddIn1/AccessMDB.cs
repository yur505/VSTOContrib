using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data.OleDb;
using System.Data;
using System.Windows.Forms;
  
namespace yxdain
{
    public class AccessHelper
    {
        private string conn_str = null;
        private OleDbConnection ole_connection = null;
        private OleDbCommand ole_command = null;
        private OleDbDataReader ole_reader = null;
        private DataTable dt = null;

        /// <summary>
        /// 构造函数
        /// </summary>
        public AccessHelper()
        {
            //conn_str = @"Provider=Microsoft.Jet.OLEDB.4.0;Data Source='" + Environment.CurrentDirectory + "\\yxdain.accdb'";
            conn_str = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" + Environment.CurrentDirectory + "\\yxdain.accdb'";

            InitDB();
        }

        private void InitDB()
        {
            ole_connection = new OleDbConnection(conn_str);//创建实例
            ole_command = new OleDbCommand();
        }

        /// <summary>
        /// 构造函数
        /// </summary>
        ///<param name="db_path">数据库路径
        public AccessHelper(string db_path)
        {
            //conn_str ="Provider=Microsoft.Jet.OLEDB.4.0;Data Source='"+ db_path + "'";
            conn_str = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" + db_path + "'";

            InitDB();
        }

        /// <summary>
        /// 转换数据格式
        /// </summary>
        ///<param name="reader">数据源
        /// <returns>数据列表</returns>
        private DataTable ConvertOleDbReaderToDataTable(ref OleDbDataReader reader)
        {
            DataTable dt_tmp = null;
            DataRow dr = null;
            int data_column_count = 0;
            int i = 0;

            data_column_count = reader.FieldCount;
            dt_tmp = BuildAndInitDataTable(data_column_count);

            if (dt_tmp == null)
            {
                return null;
            }

            while (reader.Read())
            {
                dr = dt_tmp.NewRow();

                for (i = 0; i < data_column_count; ++i)
                {
                    dr[i] = reader[i];
                }

                dt_tmp.Rows.Add(dr);
            }

            return dt_tmp;
        }

        /// <summary>
        /// 创建并初始化数据列表
        /// </summary>
        ///<param name="Field_Count">列的个数
        /// <returns>数据列表</returns>
        private DataTable BuildAndInitDataTable(int Field_Count)
        {
            DataTable dt_tmp = null;
            DataColumn dc = null;
            int i = 0;

            if (Field_Count <= 0)
            {
                return null;
            }

            dt_tmp = new DataTable();

            for (i = 0; i < Field_Count; ++i)
            {
                dc = new DataColumn(i.ToString());
                dt_tmp.Columns.Add(dc);
            }

            return dt_tmp;
        }

        /// <summary>
        /// 从数据库里面获取数据
        /// </summary>
        ///<param name="strSql">查询语句
        /// <returns>数据列表</returns>
        public DataTable GetDataTableFromDB(string strSql)
        {
            if (conn_str == null)
            {
                return null;
            }

            try
            {
                ole_connection.Open();//打开连接

                if (ole_connection.State == ConnectionState.Closed)
                {
                    return null;
                }

                ole_command.CommandText = strSql;
                ole_command.Connection = ole_connection;

                ole_reader = ole_command.ExecuteReader(CommandBehavior.Default);

                dt = ConvertOleDbReaderToDataTable(ref ole_reader);

                ole_reader.Close();
                ole_reader.Dispose();
            }
            catch (System.Exception e)
            {
                //Console.WriteLine(e.ToString());
                MessageBox.Show(e.Message);
            }
            finally
            {
                if (ole_connection.State != ConnectionState.Closed)
                {
                    ole_connection.Close();
                }
            }

            return dt;
        }

        /// <summary>
        /// 执行sql语句
        /// </summary>
        ///<param name="strSql">sql语句
        /// <returns>返回结果</returns>
        public int ExcuteSql(string strSql)
        {
            int nResult = 0;

            try
            {
                ole_connection.Open();//打开数据库连接
                if (ole_connection.State == ConnectionState.Closed)
                {
                    return nResult;
                }

                ole_command.Connection = ole_connection;
                ole_command.CommandText = strSql;

                nResult = ole_command.ExecuteNonQuery();
            }
            catch (System.Exception e)
            {
                //Console.WriteLine(e.ToString());
                MessageBox.Show(e.Message);
                return nResult;
            }
            finally
            {
                if (ole_connection.State != ConnectionState.Closed)
                {
                    ole_connection.Close();
                }
            }

            return nResult;
        }

#if false
        //显示数据表全部内容；
        private void databind1(string sqlstr)
        {
            DataTable dt = new DataTable();
            dt = achelp.GetDataTableFromDB(sqlstr);
            dataGridView1.DataSource = dt;
        }

        //添加记录；
        private void button4_Click(object sender, EventArgs e)
        {
            if (textBox1.Text == "" && textBox2.Text == "" && textBox3.Text == "" && textBox4.Text == "" && textBox5.Text == "" && textBox6.Text == "")
            {
                MessageBox.Show("没有要添加的内容", "M营销添加");
                return;
            }
            else
            {
                string sql = "insert into ycyx (fwhm,khmc,gsdq,dqpp,dqtc,dqzt) values ('" + textBox1.Text + "','" + textBox2.Text + "','" +
                  textBox3.Text + "','" + textBox4.Text + "','" + textBox5.Text + "','" + textBox6.Text + "')";
                int ret = achelp.ExcuteSql(sql);
                string sql1 = "select * from ycyx";
                databind1(sql1);
                textBox1.Text = "";
                textBox2.Text = "";
                textBox3.Text = "";
                textBox4.Text = "";
                textBox5.Text = "";
                textBox6.Text = "";
            }
        }

        //删除记录；
        private void button2_Click(object sender, EventArgs e)
        {
            if (dataGridView1.SelectedRows.Count < 1 || dataGridView1.SelectedRows[0].Cells[1].Value == null)
            {
                MessageBox.Show("没有选中行。", "M营销");
            }
            else
            {
                object oid = dataGridView1.SelectedRows[0].Cells[0].Value;
                if (DialogResult.No == MessageBox.Show("将删除第 " + (dataGridView1.CurrentCell.RowIndex + 1).ToString() + " 行，确定？", "M营销", MessageBoxButtons.YesNo))
                {
                    return;
                }
                else
                {
                    string sql = "delete from ycyx where ID=" + oid;
                    int ret = achelp.ExcuteSql(sql);
                }
                string sql1 = "select * from ycyx";
                databind1(sql1);
            }
        }

        //查询；
        private void button13_Click(object sender, EventArgs e)
        {
            if (textBox23.Text == "")
            {
                MessageBox.Show("请输入要查询的当前品牌", "M营销");
                return;
            }
            else
            {
                string sql = "select * from ycyx where dqpp='" + textBox23.Text + "'";
                DataTable dt = new System.Data.DataTable();
                dt = achelp.GetDataTableFromDB(sql);
                dataGridView1.DataSource = dt;
            }
        }

        //更新数据；
        public partial class Form3 : Form
        {
            private AccessHelper achelp;
            private int iid;

            public Form3()
            {
                InitializeComponent();
                achelp = new AccessHelper();
                iid = 0;
            }

            // 更新
            private void button1_Click(object sender, EventArgs e)
            {
                try
                {
                    //UPDATE Person SET Address = 'Zhongshan 23', City = 'Nanjing'WHERE LastName = 'Wilson'
                    string sql = "update ycyx set fwhm='" + textBox1.Text + "',khmc='" + textBox2.Text + "',gsdq='" + textBox3.Text + "',dqpp='" + textBox4.Text +
                      "',dqtc='" + textBox5.Text + "',dqzt='" + textBox6.Text + "' where ID=" + iid;


                    int ret = achelp.ExcuteSql(sql);
                    if (ret > -1)
                    {
                        this.Hide();
                        MessageBox.Show("更新成功", "M营销");
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }



            }

            private void Form3_Load(object sender, EventArgs e)
            {

            }

            public int id
            {
                get { return this.iid; }
                set { this.iid = value; }
            }


            public string Text1
            {
                get { return this.textBox1.Text; }
                set { this.textBox1.Text = value; }
            }

            public string Text2
            {
                get { return this.textBox2.Text; }
                set { this.textBox2.Text = value; }
            }

            public string Text3
            {
                get { return this.textBox3.Text; }
                set { this.textBox3.Text = value; }
            }

            public string Text4
            {
                get { return this.textBox4.Text; }
                set { this.textBox4.Text = value; }
            }

            public string Text5
            {
                get { return this.textBox5.Text; }
                set { this.textBox5.Text = value; }
            }

            public string Text6
            {
                get { return this.textBox6.Text; }
                set { this.textBox6.Text = value; }
            }

            //取消
            private void button2_Click(object sender, EventArgs e)
            {
                this.Hide();
            }
        }

        //定义变量，设置列标题；
        private void Form1_Load(object sender, EventArgs e)
        {
            achelp = new AccessHelper();
            string sql1 = "select * from ycyx";
            databind1(sql1);

            dataGridView1.Columns[0].Visible = false;
            dataGridView1.Columns[1].HeaderCell.Value = "服务号码";
            dataGridView1.Columns[2].HeaderCell.Value = "客户名称";
            dataGridView1.Columns[3].HeaderCell.Value = "归属地区";
            dataGridView1.Columns[4].HeaderCell.Value = "当前品牌";
            dataGridView1.Columns[5].HeaderCell.Value = "当前套餐";
            dataGridView1.Columns[6].HeaderCell.Value = "当前状态";
        }

        //读取要更新记录到更新窗体控件；
        private void button3_Click(object sender, EventArgs e)
        {
            if (dataGridView1.SelectedRows.Count < 1 || dataGridView1.SelectedRows[0].Cells[1].Value == null)
            {
                MessageBox.Show("没有选中行。", "M营销");
                return;
            }
            //f3.Owner = this;
            DataTable dt = new DataTable();
            object oid = dataGridView1.SelectedRows[0].Cells[0].Value;
            string sql = "select * from ycyx where ID=" + oid;
            dt = achelp.GetDataTableFromDB(sql);
            f3 = new Form3();
            f3.id = int.Parse(oid.ToString());
            //f3.id = 2;
            f3.Text1 = dt.Rows[0][1].ToString();
            f3.Text2 = dt.Rows[0][2].ToString();
            f3.Text3 = dt.Rows[0][3].ToString();
            f3.Text4 = dt.Rows[0][4].ToString();
            f3.Text5 = dt.Rows[0][5].ToString();
            f3.Text6 = dt.Rows[0][6].ToString();

            f3.ShowDialog();

        }
#endif
    }
}