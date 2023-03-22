using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.DirectoryServices;
using System.DirectoryServices.AccountManagement;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace TokuchoBugyoK2
{
    public partial class Login : Form
    {
        public Login()
        {
            InitializeComponent();
        }

        public Boolean logoff = false;
        public string ID;
        public string UserName;
        public string Roll;
        public string Busho;
        public string BushoMei;
        public string TokuchoRole;
        public string NippouRole;
        public string KanrishokuFlag;
        GlobalMethod globalMethod = new GlobalMethod();
        private string[] UserInfos;
        string uName = Environment.UserName;        // ユーザー名	
        string domain = Environment.UserDomainName; // ドメイン名	

        private void button2_Click(object sender, EventArgs e)
        {
            Boolean ErrorFlg = false;
            if (UserID.Text == "")
            {
                label3.Text = globalMethod.GetMessage("E00001", "");
                ErrorFlg = true;
            }
            if (UserPW.Text == "")
            {
                label3.Text = globalMethod.GetMessage("E00002", "");
                ErrorFlg = true;
            }

            if (!ErrorFlg)
            {
                if (UserID.Text == globalMethod.GetCommonValue1("SYSTEMMANAGER") && UserPW.Text == globalMethod.GetCommonValue2("SYSTEMMANAGER"))
                {
                    ID = "0";
                    UserName = "システム管理者";
                    Roll = "2";
                    Busho = "127950";
                    BushoMei = "システム管理者";
                    TokuchoRole = "2";
                    NippouRole = "2";
                    KanrishokuFlag = "*";

                    globalMethod.outputLogger("SystemLogin", "ユーザマスタ " + UserID.Text + " &cUSER_PASSWORD " + UserPW.Text, "", "DEBUG");
                    //各フォームへの引数
                    UserInfos = new string[] { this.ID, this.UserName, this.Busho, this.BushoMei, TokuchoRole, "", "" };

                    TopMenu form = new TopMenu();
                    form.UserInfos = this.UserInfos;
                    form.Show(this);
                    label3.Text = "";
                    UserID.Text = "";
                    UserPW.Text = "";
                    this.logoff = false;
                }
                else
                {
                    var dt = globalMethod.Check_Login(UserID.Text, UserPW.Text);

                    if (dt == null)
                    {
                        string ADDomain = globalMethod.GetCommonValue1("ACTIVE_DIRECTORY");
                        try
                        {
                            using (var context = new PrincipalContext(ContextType.Domain, ADDomain))
                            {
                                if (context.ValidateCredentials(UserID.Text, UserPW.Text, ContextOptions.Negotiate))
                                {
                                    var dt2 = globalMethod.Check_Login_Chousain("", UserID.Text);
                                    if (dt2 != null)
                                    {
                                        ID = dt2.Rows[0][0].ToString();
                                        UserName = dt2.Rows[0][1].ToString();
                                        Busho = dt2.Rows[0][2].ToString();
                                        BushoMei = dt2.Rows[0][3].ToString();
                                        TokuchoRole = dt2.Rows[0][4].ToString();

                                        globalMethod.outputLogger("LDAPLogin", "調査員マスタ： " + ID, "", "DEBUG");

                                        //各フォームへの引数
                                        UserInfos = new string[] { this.ID, this.UserName, this.Busho, this.BushoMei, TokuchoRole, "", "" };

                                        TopMenu form = new TopMenu();
                                        form.UserInfos = this.UserInfos;
                                        form.Show(this);
                                        label3.Text = "";
                                        UserID.Text = "";
                                        UserPW.Text = "";
                                        this.logoff = false;
                                    }
                                    else
                                    {
                                        label3.Text = globalMethod.GetMessage("E00005", "");
                                    }
                                }
                                else
                                {
                                    label3.Text = globalMethod.GetMessage("E00003", "");
                                }
                            }
                        }
                        catch
                        {
                            label3.Text = globalMethod.GetMessage("E00003", "");
                        }
                    }
                    else
                    {
                        ID = dt.Rows[0][0].ToString();
                        //UserName = dt.Rows[0][1].ToString();
                        Roll = dt.Rows[0][2].ToString();

                        globalMethod.outputLogger("UserLogin", "ユーザマスタ検索HIT=デバッグ用ユーザ有り " + UserID.Text, "", "DEBUG");

                        var dt2 = globalMethod.Check_Login_Chousain(ID, "");

                        if (dt2 == null)
                        {
                            label3.Text = globalMethod.GetMessage("E00005", "");
                        }
                        else
                        {
                            UserName = dt2.Rows[0][1].ToString();
                            Busho = dt2.Rows[0][2].ToString();
                            BushoMei = dt2.Rows[0][3].ToString();
                            TokuchoRole = dt2.Rows[0][4].ToString();

                            globalMethod.outputLogger("UserLogin", "L252 ChousainMei  " + UserName + " GyoumuBushoCD " + Busho, "", "DEBUG");

                            //各フォームへの引数
                            UserInfos = new string[] { this.ID, this.UserName, this.Busho, this.BushoMei, TokuchoRole, "", "" };

                            TopMenu form = new TopMenu();
                            form.UserInfos = this.UserInfos;
                            form.Show(this);

                            label3.Text = "";
                            UserID.Text = "";
                            UserPW.Text = "";
                            this.logoff = false;
                        }
                    }
                }
            }
        }

        private void Login_Load(object sender, EventArgs e)
        {
        }

        private void GetADUser()
        {
            string path = "WinNT://" + domain + "/" + uName;

            string dName;
            // ADに問い合わせ	
            try
            {
                using (DirectoryEntry dirEnt = new DirectoryEntry(path))
                {
                    // ネーム	
                    dName = dirEnt.Properties["Name"].Value.ToString();
                }

                if (dName != null && dName != "")
                {
                    var dt = globalMethod.Check_Login_Chousain("", uName);
                    if (dt != null)
                    {
                        ID = dt.Rows[0][0].ToString();
                        UserName = dt.Rows[0][1].ToString();
                        Busho = dt.Rows[0][2].ToString();
                        BushoMei = dt.Rows[0][3].ToString();
                        TokuchoRole = dt.Rows[0][4].ToString();


                        globalMethod.outputLogger("LDAPLogin", "調査員マスタ： " + ID, "", "DEBUG");

                        //各フォームへの引数
                        UserInfos = new string[] { this.ID, this.UserName, this.Busho, this.BushoMei, TokuchoRole, "", "" };

                        TopMenu form = new TopMenu();
                        form.UserInfos = this.UserInfos;
                        form.Show(this);
                    }
                    else
                    {
                        this.Show();
                    }
                }
                else
                {
                    this.Show();
                }
            }
            catch (Exception)
            {
                this.Show();
            }

        }

        private void Login_Shown(object sender, EventArgs e)
        {
            if (globalMethod.Check_DB())
            {
                if (globalMethod.GetCommonValue1("ACTIVE_DIRECTORY") == domain)
                { 
                    if (!logoff)
                    {
                        GetADUser();
                    }
                }
            }
            else
            {
                Application.Exit();
            }
        }
    }
}
