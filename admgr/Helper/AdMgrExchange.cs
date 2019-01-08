using admgr.EfModel;
using System;
using System.Collections.Generic;
using System.DirectoryServices;
using System.Linq;
using System.Management.Automation;
using System.Management.Automation.Runspaces;
using System.Security;
using System.Text;
using System.Web;

namespace ADmgr.Helper
{
    public class AdMgrExchange
    {
        readonly DirectoryEntry _root;
        readonly string _exIp;
        readonly string _exadmin;
        readonly string _expwd;

        public AdMgrExchange(DirectoryEntry root)
        {
            _root = root;
            AdmgrModel model = new AdmgrModel();
            var admgrSystem = model.AdmgrSystems.FirstOrDefault(a => a.AdmgrSystemId == 1);
            _exIp = admgrSystem.ExchangIp;
            _exadmin = admgrSystem.ExchangeUserName;
            _expwd = admgrSystem.ExchangePwd;
            model.Dispose();
        }

        #region 注册邮箱

        public string EnableMail(string identity)
        {
            var message = "";
            var success = "";

            #region 取得Exchange所需要的数据
            var array = identity.Split('/');
            var displayName = array[array.Length - 1];
            if (!displayName.Contains(".") || !displayName.Contains("-"))
            {
                return "邮件开启失败！";
            }
            var alias = displayName.Substring(0, displayName.IndexOf('.'));
            //var ouDeptPath = identity.Substring(0, identity.Length - displayName.Length);
            var updateName = identity.Substring(identity.IndexOf('/') + 1, identity.LastIndexOf('/') - identity.IndexOf('/') - 1).Replace('/', '\\');
            var ouCompanyName = array[array.Length - 3];
            //根据公司的名称，在数据库中找到该公司所对应的数据库名称，并赋值给dataBase变量
            var emailData = (new AdmgrModel()).EmailDatas.FirstOrDefault(a => a.CompanyName == ouCompanyName);
            string dataBase = emailData.EmailStoreDb;
            #endregion
            var search = new DirectorySearcher(_root);
            search.Filter = "(name=" + displayName + ")";
            //search.PropertiesToLoad.Add("cn");
            var result = search.FindOne();
            //确认AD使用者是否存在
            if (result == null)
            {
                message = "该AD域用户不存在！";
            }
            else
            {
                //判断使用者mail栏是否有值
                var user = result.GetDirectoryEntry();

                //工号（在域里字段是名）
                if (user.Properties.Contains("homeMDB") && user.Properties.Contains("homeMTA"))
                {
                    message = "该AD域用户邮件已开启！";
                }
                else if (dataBase == null)
                {
                    message = "邮箱数据库未定义！";
                }
                else
                {
                    try
                    {
                        // 生成一个连接类型，传入exchange服务器IP、将要使用的Scheme以及管理员凭据
                        var connectionInfo = ExchangeScriptInit();
                        // 创建一个命令空间，创建管线通道，传入要运行的powershell命令，执行命令
                        using (var rs = RunspaceFactory.CreateRunspace(connectionInfo))
                        {
                            var psh = PowerShell.Create();
                            psh.Runspace = rs;
                            rs.Open();

                            var sScript = "Enable-Mailbox -identity  '" + identity + "' -Alias '" + alias + "' -Database '" + dataBase + "'";
                            psh.Commands.AddScript(sScript);
                            //var sScript2 = "update-AddressList -identity '" + updateName + "'";
                            //psh.Commands.AddScript(sScript2);
                            var psresults = psh.Invoke();
                            if (psresults == null)
                            {
                                message = "邮件开启失败！";
                            }
                            if (psh.Streams.Error.Count > 0)
                            {
                                var strbmessage = new StringBuilder();
                                foreach (var err in psh.Streams.Error)
                                {
                                    strbmessage.AppendLine(err.ToString());
                                }
                                message = strbmessage.ToString();
                            }
                            else
                            {
                                success = "1";
                            }
                            rs.Close();
                            psh.Runspace.Close();
                        }
                    }
                    catch (Exception e)
                    {

                        return e.Message + "：邮件开启失败！";
                    }
                    finally
                    {

                    }
                    if (success == "1")
                    {
                        message = "邮件开启成功！";
                    }
                    else
                    {
                        message = "邮件开启失败！" + message;
                    }
                }
            }

            return message;
        }

        #endregion

        #region 在全局地址列表上添加公司或者部门
        // adCompanyName的名称是全局名称 test.com/xd-ad/集团公司
        //AllRecipients是所有收件人类型的意思，RecipientContainer收件人容器应以为AD中的全路径，ConditionalCompany 条件公司名域与name、displayName相同
        public bool AddGolbalCompanyAddress(string adCompanyName)
        {
            string companyName = adCompanyName.Substring(adCompanyName.LastIndexOf('/') + 1);
            try
            {
                // 生成一个连接类型，传入exchange服务器IP、将要使用的Scheme以及管理员凭据
                var connectionInfo = ExchangeScriptInit();

                using (var rs = RunspaceFactory.CreateRunspace(connectionInfo))
                {
                    var psh = PowerShell.Create();
                    psh.Runspace = rs;
                    rs.Open();
                    var sScript = "New-AddressList -name '" + companyName + "' -recipientContainer '" + adCompanyName + "' -IncludedRecipients 'AllRecipients' -ConditionalCompany '" + companyName + "' -Container '\\' -displayName '" + companyName + "'";

                    //string sScript = "Enable-Mailbox -identity  'xd-ad\\" + alias + "' -Alias '" + email + "' -Database 'jt-db'";
                    psh.Commands.AddScript(sScript);
                    var sScript2 = "Update-AddressList -identity '" + companyName + "'";
                    psh.Commands.AddScript(sScript2);
                    var psresults = psh.Invoke();
                    if (psresults == null)
                    {
                        return false;
                    }
                    var strbmessage = new StringBuilder();
                    if (psh.Streams.Error.Count > 0)
                    {
                        foreach (var err in psh.Streams.Error)
                        {
                            strbmessage.AppendLine(err.ToString());
                        }
                        return false;
                    }

                    rs.Close();
                    psh.Runspace.Close();
                }
            }
            catch (Exception e)
            {
                throw new Exception(e.Message);
            }
            return true;
        }

        #endregion

        #region 增加通讯组元素
        //不需要更细地址列表，因为增加通讯组元素的时候，肯定会增加地址列表，如果OU下有不同于OU名的组，这个情况就不对了
        //ouName是组名，alias是组的别名，特别注意，ouName是6种类型中的一种
        //eg:ouName= test.com/xd-ad/bim公司/bim公司，alias=1234，alias等于ouName名字中.符号前很长的数字
        //
        public bool AddDistributionGroup(string identity)
        {
            try
            {
                // 生成一个连接类型，传入exchange服务器IP、将要使用的Scheme以及管理员凭据
                var connectionInfo = ExchangeScriptInit();
                var array = identity.Split('/');
                var displayName = array[array.Length - 1];
                var alias = displayName.Substring(0, displayName.IndexOf('.'));
                var groupIdentity = identity + "/" + displayName;
                using (var rs = RunspaceFactory.CreateRunspace(connectionInfo))
                {
                    var psh = PowerShell.Create();
                    psh.Runspace = rs;
                    rs.Open();
                    var sScript = "Enable-DistributionGroup -identity '" + groupIdentity + "' -Alias '" + alias + "'";
                    psh.Commands.AddScript(sScript);
                    var psresults = psh.Invoke();
                    if (psresults == null)
                    {
                        return false;
                    }
                    var strbmessage = new StringBuilder();
                    if (psh.Streams.Error.Count > 0)
                    {
                        foreach (var err in psh.Streams.Error)
                        {
                            strbmessage.AppendLine(err.ToString());
                        }
                        return false;
                    }

                    rs.Close();
                    psh.Runspace.Close();
                }
            }
            catch (Exception e)
            {
                throw new Exception(e.Message);
            }
            return true;
        }
        #endregion

        #region 重命名部门的时候，修改通讯组，并修改地址列表元素
        public bool ModifyDistributionGroupAndModifyAddressList_dept(string identity, string oldIdentity)
        {
            // 生成一个连接类型，传入exchange服务器IP、将要使用的Scheme以及管理员凭据
            var connectionInfo = ExchangeScriptInit();

            var array = identity.Split('/');
            //别名alias约定成.前的一串数字
            var displayName = array[array.Length - 1];
            var alias = displayName.Substring(0, displayName.IndexOf('.'));
            var identityGroupName = identity + "/" + displayName;
            var oldIdentityAddressName = oldIdentity.Substring(oldIdentity.IndexOf('/') + 1, oldIdentity.Length - oldIdentity.IndexOf('/') - 1).Replace('/', '\\');
            var identityAddressName = identity.Substring(identity.IndexOf('/') + 1, identity.Length - identity.IndexOf('/') - 1).Replace('/', '\\');
            // OU为部门组的时候，identity类似 test.com/xd-ad/集团公司/办公室/办公室 这种名称，而recipientContainer需要的是组的名称
            // 去掉最后一个/后面的内容，就是recipientContainer需要的内容
            var recipientContainer = identity.Substring(0, identity.Length - displayName.Length - 1);

            //
            try
            {
                using (var rs = RunspaceFactory.CreateRunspace(connectionInfo))
                {
                    var psh = PowerShell.Create();
                    psh.Runspace = rs;
                    rs.Open();
                    var sScript = "Set-DistributionGroup -DisplayName '" + displayName + "'  -Identity '" + identityGroupName + "'";
                    //Set-DistributionGroup -DisplayName '0100-01011.集团公司领导班子' -Name '0100-01011.集团公司领导班子' -Identity 'test.com/xd-ad/0100-01001.集团公司/0100-01011.集团公司领导l班子/0100-01011.集团公司领导班子'
                    psh.Commands.AddScript(sScript);
                    // 这个oldIdentity就是地址列表中ou以前的名称
                    var sScript1 = "set-AddressList -DisplayName '" + displayName + "' -Name '" + displayName + "' -IncludedRecipients 'AllRecipients' -ConditionalDepartment '" + displayName + "' -RecipientContainer '" + recipientContainer + "' -identity '\\" + oldIdentityAddressName + "'";
                    // set-AddressList -DisplayName '0100-01011.集团公司领导班子' -Name '0100-01011.集团公司领导班子' -IncludedRecipients 'AllRecipients' -ConditionalDepartment '0100-01011.集团公司领导班子' -RecipientContainer 'test.com/xd-ad临时员工/0100-01001.集团公司/0100-01011.集团公司领导l班子' -Identity '\xd-ad临时员工\0100-01001.集团公司\0100-01011.集团公司领导'
                    psh.Commands.AddScript(sScript1);
                    // 这里更改部门上一级的公司OU存在争议，可改可不改暂时不清楚，暂时不写
                    var sScript2 = "update-AddressList -identity '\\" + identityAddressName + "'";
                    //updateName "xd-ad被禁账户\\0100-01001.集团公司\\0100-01011.集团公司领导l班子"
                    psh.Commands.AddScript(sScript2);
                    var psresults = psh.Invoke();
                    if (psresults == null)
                    {
                        return false;
                    }
                    var strbmessage = new StringBuilder();
                    if (psh.Streams.Error.Count > 0)
                    {
                        foreach (var err in psh.Streams.Error)
                        {
                            strbmessage.AppendLine(err.ToString());
                        }
                        return false;
                    }
                    rs.Close();
                    psh.Runspace.Close();
                    return true;
                }
            }
            catch (Exception e)
            {
                throw new Exception(e.Message);
            }

        }
        #endregion

        #region 重命名公司的时候，修改通讯组，并修改地址列表元素
        /// <summary>
        /// 返回一个是否执行成功的布尔值
        /// </summary>
        /// <param name="alias">部门组的别名</param>
        /// <param name="identity">部门组的全路径</param>
        /// <returns></returns>
        public bool ModifyDistributionGroupAndModifyAddressList_company(string identity)
        {
            // 生成一个连接类型，传入exchange服务器IP、将要使用的Scheme以及管理员凭据
            var connectionInfo = ExchangeScriptInit();
            var array = identity.Split('/');
            var displayName = array[array.Length - 1];
            var alias = displayName.Substring(0, displayName.IndexOf('.'));
            var name = array[array.Length - 1];
            var updateName = identity.Substring(identity.IndexOf('/') + 1).Replace('/', '\\');
            //筛选名称
            var conditionnalName = array[array.Length - 1];
            // OU为部门组的时候，identity类似 test.com/xd-ad/集团公司/办公室/办公室 这种名称，而recipientContainer需要的是组的名称
            // 去掉最后一个/后面的内容，就是recipientContainer需要的内容
            var recipientContainer = identity.Substring(0, identity.Length - displayName.Length - 1);
            try
            {
                using (var rs = RunspaceFactory.CreateRunspace(connectionInfo))
                {
                    var psh = PowerShell.Create();
                    psh.Runspace = rs;
                    rs.Open();
                    var sScript = "Set-DistributionGroup -Alias '" + alias + "' -DisplayName '" + displayName + "' -Identity '" + identity + "'";
                    //var sScript = "Set-DistributionGroup -Alias '"+ alias + """' -DisplayName '炊事房1' -Identity 'test.com/xd-ad/集团公司/炊事房/炊事房'";
                    psh.Commands.AddScript(sScript);
                    var sScript1 = "set-AddressList -displayName '" + displayName + "' -name '" + name + "' -IncludedRecipients 'AllRecipients' -ConditionalCompany  '" + conditionnalName + "' -recipientContainer '" + recipientContainer + "' -identity '\\" + identity + "'";
                    psh.Commands.AddScript(sScript1);
                    var sScript2 = "update-AddressList -identity '\\" + updateName + "'";
                    psh.Commands.AddScript(sScript2);
                    var psresults = psh.Invoke();
                    if (psresults == null)
                    {
                        return false;
                    }
                    var strbmessage = new StringBuilder();
                    if (psh.Streams.Error.Count > 0)
                    {
                        foreach (var err in psh.Streams.Error)
                        {
                            strbmessage.AppendLine(err.ToString());
                        }
                        return false;
                    }
                    rs.Close();
                    psh.Runspace.Close();
                    return true;
                }
            }
            catch (Exception e)
            {
                throw new Exception(e.Message);
            }

        }
        #endregion

        #region 新增部门组的时候，新增通讯组，新增地址列表元素
        /// <summary>
        /// 返回一个是否执行成功的布尔值
        /// </summary>
        /// <param name="alias">部门组的别名</param>
        /// <param name="identity">部门组的全路径</param>
        /// <returns></returns>
        public bool AddDistributionGroupAndAddAddressList_dept(string identity)
        {
            try
            {
                // 生成一个连接类型，传入exchange服务器IP、将要使用的Scheme以及管理员凭据
                var connectionInfo = ExchangeScriptInit();
                var array = identity.Split('/');
                var displayName = array[array.Length - 1];
                var alias = displayName.Substring(0, displayName.IndexOf('.'));
                var name = array[array.Length - 1];
                //筛选名称
                var conditionnalName = array[array.Length - 1];
                //container是部门上一级：公司的名称

                var recipientContainer = identity.Substring(0, identity.Length - displayName.Length - 1);
                var updateName = identity.Substring(identity.IndexOf('/') + 1).Replace('/', '\\');
                var temp = identity.Substring(identity.IndexOf('/') + 1);
                var container = temp.Substring(0, temp.LastIndexOf('/')).Replace('/', '\\');
                var groupIdentity = identity + "/" + displayName;
                using (var rs = RunspaceFactory.CreateRunspace(connectionInfo))
                {
                    var psh = PowerShell.Create();
                    psh.Runspace = rs;
                    rs.Open();
                    var sScript = "Enable-DistributionGroup -identity '" + groupIdentity + "' -Alias '" + alias + "'";
                    psh.Commands.AddScript(sScript);
                    var sScript1 = "new-AddressList -Name '" + name + "' -RecipientContainer '" + identity + "' -IncludedRecipients 'AllRecipients' -ConditionalDepartment '" + conditionnalName + "' -Container '\\" + container + "' -DisplayName '" + displayName + "'";
                    //new-AddressList -Name 'test5' -RecipientContainer 'test.com/xd-ad/集团公司/党委' -IncludedRecipients 'AllRecipients' -ConditionalDepartment '党委' -Container '\\集团公司' -DisplayName 'test5'
                    psh.Commands.AddScript(sScript1);
                    var sScript2 = "update-AddressList -identity '\\" + updateName + "'";
                    psh.Commands.AddScript(sScript2);
                    var psresults = psh.Invoke();
                    if (psresults == null)
                    {
                        return false;
                    }
                    var strbmessage = new StringBuilder();
                    if (psh.Streams.Error.Count > 0)
                    {
                        foreach (var err in psh.Streams.Error)
                        {
                            strbmessage.AppendLine(err.ToString());
                        }
                        return false;
                    }

                    rs.Close();
                    psh.Runspace.Close();
                }
            }
            catch (Exception e)
            {
                throw new Exception(e.Message);
            }
            return true;
        }
        #endregion

        #region 新增公司组的时候，新增通讯组，新增地址列表元素
        /// <summary>
        /// 返回一个是否执行成功的布尔值
        /// </summary>
        /// <param name="alias">公司的别名</param>
        /// <param name="identity">公司组的全路径</param>
        /// <returns></returns>
        public bool AddDistributionGroupAndAddAddressList_company(string identity)
        {
            try
            {
                // 生成一个连接类型，传入exchange服务器IP、将要使用的Scheme以及管理员凭据
                var connectionInfo = ExchangeScriptInit();
                var array = identity.Split('/');
                var displayName = array[array.Length - 1];
                var alias = displayName.Substring(0, displayName.IndexOf('.'));
                var name = array[array.Length - 1];
                //筛选名称
                var conditionnalName = array[array.Length - 1];
                //container是部门上一级：帐号类型的名称
                //var container = array[array.Length - 2];
                var updateName = identity.Substring(identity.IndexOf('/') + 1).Replace('/', '\\');
                var temp = identity.Substring(identity.IndexOf('/') + 1);
                var container = temp.Substring(0, temp.LastIndexOf('/')).Replace('/', '\\');
                var groupIdentity = identity + "/" + displayName;
                using (var rs = RunspaceFactory.CreateRunspace(connectionInfo))
                {
                    var psh = PowerShell.Create();
                    psh.Runspace = rs;
                    rs.Open();
                    var sScript = "Enable-DistributionGroup -identity '" + groupIdentity + "' -Alias '" + alias + "'";
                    psh.Commands.AddScript(sScript);
                    var sScript1 = "new-AddressList -Name '" + name + "' -RecipientContainer '" + identity + "' -IncludedRecipients 'AllRecipients' -ConditionalCompany  '" + conditionnalName + "' -Container '\\" + container + "' -DisplayName '" + displayName + "'";
                    //new-AddressList -Name 'test5' -RecipientContainer 'test.com/xd-ad/集团公司/党委' -IncludedRecipients 'AllRecipients' -ConditionalDepartment '党委' -Container '\\集团公司' -DisplayName 'test5'
                    psh.Commands.AddScript(sScript1);

                    var sScript2 = "update-AddressList -identity '\\" + updateName + "'";
                    psh.Commands.AddScript(sScript2);
                    var psresults = psh.Invoke();
                    if (psresults == null)
                    {
                        return false;
                    }
                    var strbmessage = new StringBuilder();
                    if (psh.Streams.Error.Count > 0)
                    {
                        foreach (var err in psh.Streams.Error)
                        {
                            strbmessage.AppendLine(err.ToString());
                        }
                        return false;
                    }

                    rs.Close();
                    psh.Runspace.Close();
                }
            }
            catch (Exception e)
            {
                throw new Exception(e.Message);
            }
            return true;
        }
        #endregion 

        #region 修改通讯组元素
        public bool ModifyDistributionGroup(string identity)
        {
            try
            {
                // 生成一个连接类型，传入exchange服务器IP、将要使用的Scheme以及管理员凭据
                var connectionInfo = ExchangeScriptInit();
                //下面的这行未测试
                var displayName = identity.Split('/')[identity.Split('/').Length - 1];
                var alias = displayName.Substring(0, displayName.IndexOf('.') - 1);
                using (var rs = RunspaceFactory.CreateRunspace(connectionInfo))
                {
                    var psh = PowerShell.Create();
                    psh.Runspace = rs;
                    rs.Open();
                    //var sScript = "Enable-DistributionGroup -identity '" + ouName + "' -Alias '" + alias + "'";
                    var sScript = "Set-DistributionGroup -Alias '" + alias + "' -DisplayName '" + displayName + "' -Identity '" + identity + "'";
                    //var sScript = "Set-DistributionGroup -Alias '"+ alias + """' -DisplayName '炊事房1' -Identity 'test.com/xd-ad/集团公司/炊事房/炊事房'";
                    psh.Commands.AddScript(sScript);
                    var psresults = psh.Invoke();
                    if (psresults == null)
                    {
                        return false;
                    }
                    var strbmessage = new StringBuilder();
                    if (psh.Streams.Error.Count > 0)
                    {
                        foreach (var err in psh.Streams.Error)
                        {
                            strbmessage.AppendLine(err.ToString());
                        }
                        return false;
                    }

                    rs.Close();
                    psh.Runspace.Close();
                }
            }
            catch (Exception e)
            {
                throw new Exception(e.Message);
            }
            return true;
        }
        #endregion

        #region 设定用户的邮箱容量
        /// <summary>
        /// 返回是否操作成功的布尔值
        /// </summary>
        /// <param name="adName"></param>
        /// <param name="issusWarningQuota">达到该限度时发出警告</param>
        /// <param name="prohibitSendQuota">达到该限度时禁止发送</param>
        /// <param name="prohibitSendReceiveQuota">达到该限度时禁止发送和接收</param>
        /// <returns></returns>
        public bool AddDistributionGroup(string adName, int issusWarningQuota, int prohibitSendQuota, int prohibitSendReceiveQuota)
        {
            try
            {
                // 生成一个连接类型，传入exchange服务器IP、将要使用的Scheme以及管理员凭据
                var connectionInfo = ExchangeScriptInit();

                using (var rs = RunspaceFactory.CreateRunspace(connectionInfo))
                {
                    var psh = PowerShell.Create();
                    psh.Runspace = rs;
                    rs.Open();
                    //var sScript = "Enable-DistributionGroup -identity '" + ouName + "' -Alias '" + alias + "'";
                    StringBuilder sScript = new StringBuilder();
                    sScript.Append("Set-Mailbox ");
                    sScript.Append("-UseDatabaseQuotaDefaults $false ");
                    sScript.Append("-IssueWarningQuota '" + issusWarningQuota + "MB' ");
                    sScript.Append("-ProhibitSendQuota '" + prohibitSendQuota + "MB'");
                    sScript.Append("-ProhibitSendReceiveQuota '" + prohibitSendReceiveQuota + "MB' ");
                    sScript.Append("-UseDatabaseRetentionDefaults $true -identity  '" + adName + "'");
                    var ttt = sScript.ToString();
                    psh.Commands.AddScript(ttt);
                    var psresults = psh.Invoke();
                    if (psresults == null)
                    {
                        return false;
                    }
                    var strbmessage = new StringBuilder();
                    if (psh.Streams.Error.Count > 0)
                    {
                        foreach (var err in psh.Streams.Error)
                        {
                            strbmessage.AppendLine(err.ToString());
                        }
                        var tt = strbmessage.ToString();
                        return false;
                    }

                    rs.Close();
                    psh.Runspace.Close();
                }
            }
            catch (Exception e)
            {
                throw new Exception(e.Message);
            }
            return true;
        }
        #endregion

        #region 自定义邮件答复地址
        /// <summary>
        /// 现在仅能多创建一个额外的邮件地址，自动采用默认的邮件地址，这里并没有自动设置为答复地址
        /// </summary>
        /// <param name="originalEmail">创建AD账户时创建的邮件地址</param>
        /// <param name="adName">AD用户名称</param>
        /// <param name="suffix">新的邮箱后缀</param>
        /// <returns></returns>
        public bool CustomEmailAddress(string originalEmail, string adName, string suffix)
        {
            string name = adName.Substring(adName.LastIndexOf('/') + 1);
            try
            {
                // 生成一个连接类型，传入exchange服务器IP、将要使用的Scheme以及管理员凭据
                var connectionInfo = ExchangeScriptInit();

                using (var rs = RunspaceFactory.CreateRunspace(connectionInfo))
                {
                    var psh = PowerShell.Create();
                    psh.Runspace = rs;
                    rs.Open();
                    //var sScript = "Enable-DistributionGroup -identity '" + ouName + "' -Alias '" + alias + "'";
                    StringBuilder sScript = new StringBuilder();
                    sScript.Append("Set-Mailbox -EmailAddressPolicyEnabled $true ");
                    sScript.Append("-EmailAddresses 'SMTP:'" + name + "'@'" + suffix + "','" + originalEmail + "'");
                    sScript.Append("-identity '" + adName + "'");
                    var ttt = sScript.ToString();
                    psh.Commands.AddScript(ttt);
                    var psresults = psh.Invoke();
                    if (psresults == null)
                    {
                        return false;
                    }
                    var strbmessage = new StringBuilder();
                    if (psh.Streams.Error.Count > 0)
                    {
                        foreach (var err in psh.Streams.Error)
                        {
                            strbmessage.AppendLine(err.ToString());
                        }
                        var tt = strbmessage.ToString();
                        return false;
                    }

                    rs.Close();
                    psh.Runspace.Close();
                }
            }
            catch (Exception e)
            {
                throw new Exception(e.Message);
            }
            return true;
        }
        #endregion

        #region 返回相关用户的邮箱地址
        //返回用户的邮箱地址信息
        public string GetEmailadd(string userName)
        {
            var add = "";
            try
            {
                // 生成一个连接类型，传入exchange服务器IP、将要使用的Scheme以及管理员凭据
                var connectionInfo = ExchangeScriptInit();

                using (var rs = RunspaceFactory.CreateRunspace(connectionInfo))
                {
                    var psh = PowerShell.Create();
                    psh.Runspace = rs;
                    rs.Open();

                    //string sScript = "get-mailbox -identity '" + userName + "' | fl name, emailaddresses ";
                    var sScript = "Get-Mailbox -identity '" + userName + "' | select-object emailaddresses ";
                    psh.Commands.AddScript(sScript);

                    var psresults = psh.Invoke();

                    var strbmessage = new StringBuilder();
                    if (psh.Streams.Error.Count > 0)
                    {
                        foreach (var err in psh.Streams.Error)
                        {
                            strbmessage.AppendLine(err.ToString());
                        }
                        return null;
                    }
                    foreach (var obj in psresults)
                    {
                        //Console.WriteLine();
                        add = obj.Members["emailaddresses"].Value.ToString();
                    }
                    //add = psh.Streams.Error[1].ToString();

                    rs.Close();
                    psh.Runspace.Close();
                }
            }
            catch (Exception e)
            {
                throw new Exception(e.Message);
            }
            return add;
        }

        #endregion

        #region 新增地址列表
        /// <summary>
        /// AD中修改OU后，需要在exchange地址列表中同步修改对应的OU，并更新
        /// </summary>
        /// <param name="displayName">显示名称</param>
        /// <param name="name">名称</param>
        /// <param name="recipientContainer">容器的名称，需要写全OU所在的路径，eg：tets.com/xd-ad/集团公司/邮件公司</param>
        /// <returns>返回是否成功的布尔值</returns>
        public bool AddAddressList(string displayName, string name, string recipientContainer)
        {
            #region

            //RunspaceConfiguration runConf = RunspaceConfiguration.Create();
            //PSSnapInException pssException = null;
            //Runspace runsp = null;
            //Pipeline pipeline = null;
            //try
            //{
            //   PSSnapInInfo info = runConf.AddPSSnapIn("Microsoft.Exchange.Management.PowerShell.Admin", out pssException);
            //   runsp = RunspaceFactory.CreateRunspace(runConf);
            //   System.Threading.Thread.Sleep(20000);

            //   runsp.Open();
            //   pipeline = runsp.CreatePipeline();
            //   Command command = new Command("Update-AddressList"); //powershell指令
            //   command.Parameters.Add("identity", @"\" + addresslistname2);
            //   pipeline.Commands.Add(command);
            //   pipeline.Invoke();
            //   return "人员调动成功！";
            //}
            //catch (Exception ex)
            //{
            //    Debug.WriteLine(ex.Message);
            //    return "人员调动失败！";
            //}
            //finally
            //{
            //    pipeline.Dispose();
            //    runsp.Close();
            //}

            #endregion

            try
            {
                // 生成一个连接类型，传入exchange服务器IP、将要使用的Scheme以及管理员凭据
                var connectionInfo = ExchangeScriptInit();
                var array = recipientContainer.Split('/');
                string container = array[array.Length - 2];
                string conditionnalName = array[array.Length - 1];

                using (var rs = RunspaceFactory.CreateRunspace(connectionInfo))
                {
                    var psh = PowerShell.Create();
                    psh.Runspace = rs;
                    rs.Open();
                    //var sScript = "new-AddressList -Name '" + name + "' -RecipientContainer '" + recipientContainer + "' -IncludedRecipients 'AllRecipients' -ConditionalDepartment '" + conditionnalName + "' -Container '\\"+conditionnalName+"' -DisplayName '"+displayName+"'";
                    var sScript = "new-AddressList -Name '" + name + "' -RecipientContainer '" + recipientContainer + "' -IncludedRecipients 'AllRecipients' -ConditionalDepartment '" + conditionnalName + "' -Container '\\" + container + "' -DisplayName '" + displayName + "'";
                    //new-AddressList -Name 'test5' -RecipientContainer 'test.com/xd-ad/集团公司/党委' -IncludedRecipients 'AllRecipients' -ConditionalDepartment '党委' -Container '\\集团公司' -DisplayName 'test5'
                    psh.Commands.AddScript(sScript);
                    var sScript1 = "update-AddressList -identity '\\" + name + "'";
                    psh.Commands.AddScript(sScript1);
                    psh.Invoke();
                    rs.Close();
                    psh.Runspace.Close();
                    return true;
                }
            }
            catch (Exception e)
            {
                return false;
            }
        }
        #endregion 

        #region 修改地址列表
        /// <summary>
        /// AD中修改OU后，需要在exchange地址列表中同步修改对应的OU，并更新
        /// </summary>
        /// <param name="displayName">显示名称</param>
        /// <param name="name">名称</param>
        /// <param name="conditionnalName">条件名称，这里指的OU的name</param>
        /// <param name="recipientContainer">容器的名称，需要写全，eg：tets.com/xd-ad/集团公司/邮件公司</param>
        /// <param name="identity">当前对象的名称，eg："\test"</param>
        /// <returns>返回是否成功的布尔值</returns>
        public bool ModifyAddressList(string displayName, string name, string conditionnalName, string recipientContainer, string identity)
        {
            #region

            //RunspaceConfiguration runConf = RunspaceConfiguration.Create();
            //PSSnapInException pssException = null;
            //Runspace runsp = null;
            //Pipeline pipeline = null;
            //try
            //{
            //   PSSnapInInfo info = runConf.AddPSSnapIn("Microsoft.Exchange.Management.PowerShell.Admin", out pssException);
            //   runsp = RunspaceFactory.CreateRunspace(runConf);
            //   System.Threading.Thread.Sleep(20000);

            //   runsp.Open();
            //   pipeline = runsp.CreatePipeline();
            //   Command command = new Command("Update-AddressList"); //powershell指令
            //   command.Parameters.Add("identity", @"\" + addresslistname2);
            //   pipeline.Commands.Add(command);
            //   pipeline.Invoke();
            //   return "人员调动成功！";
            //}
            //catch (Exception ex)
            //{
            //    Debug.WriteLine(ex.Message);
            //    return "人员调动失败！";
            //}
            //finally
            //{
            //    pipeline.Dispose();
            //    runsp.Close();
            //}

            #endregion

            try
            {
                // 生成一个连接类型，传入exchange服务器IP、将要使用的Scheme以及管理员凭据
                var connectionInfo = ExchangeScriptInit();

                using (var rs = RunspaceFactory.CreateRunspace(connectionInfo))
                {
                    var psh = PowerShell.Create();
                    psh.Runspace = rs;
                    rs.Open();
                    var sScript = "set-AddressList -displayName '" + displayName + "' -name '" + name + "' -IncludedRecipients 'AllRecipients' -ConditionalDepartment '" + conditionnalName + "' -recipientContainer '" + recipientContainer + "' -identity '\\" + identity + "'";
                    psh.Commands.AddScript(sScript);
                    var sScript1 = "update-AddressList -identity '\\" + name + "'";
                    psh.Commands.AddScript(sScript1);
                    psh.Invoke();
                    rs.Close();
                    psh.Runspace.Close();
                    return true;
                }
            }
            catch (Exception)
            {
                return false;
            }
        }
        #endregion

        #region 更新地址列表

        //更新地址列表的操作
        public bool UpdateAddressListAfterOperate(string addresslistname)
        {
            #region

            //RunspaceConfiguration runConf = RunspaceConfiguration.Create();
            //PSSnapInException pssException = null;
            //Runspace runsp = null;
            //Pipeline pipeline = null;
            //try
            //{
            //   PSSnapInInfo info = runConf.AddPSSnapIn("Microsoft.Exchange.Management.PowerShell.Admin", out pssException);
            //   runsp = RunspaceFactory.CreateRunspace(runConf);
            //   System.Threading.Thread.Sleep(20000);

            //   runsp.Open();
            //   pipeline = runsp.CreatePipeline();
            //   Command command = new Command("Update-AddressList"); //powershell指令
            //   command.Parameters.Add("identity", @"\" + addresslistname2);
            //   pipeline.Commands.Add(command);
            //   pipeline.Invoke();
            //   return "人员调动成功！";
            //}
            //catch (Exception ex)
            //{
            //    Debug.WriteLine(ex.Message);
            //    return "人员调动失败！";
            //}
            //finally
            //{
            //    pipeline.Dispose();
            //    runsp.Close();
            //}

            #endregion

            try
            {
                // 生成一个连接类型，传入exchange服务器IP、将要使用的Scheme以及管理员凭据
                var connectionInfo = ExchangeScriptInit();

                using (var rs = RunspaceFactory.CreateRunspace(connectionInfo))
                {
                    var psh = PowerShell.Create();
                    psh.Runspace = rs;
                    rs.Open();
                    var sScript = "Update-AddressList -identity '" + addresslistname + "'";
                    psh.Commands.AddScript(sScript);
                    psh.Invoke();
                    rs.Close();
                    psh.Runspace.Close();
                    return true;
                }
            }
            catch (Exception)
            {
                return false;
            }
        }

        #endregion

        #region 设置用户为不显示，并刷新地址列表
        //更新地址列表的操作
        public bool SetUserNoShow(string identity)
        {
            #region

            //RunspaceConfiguration runConf = RunspaceConfiguration.Create();
            //PSSnapInException pssException = null;
            //Runspace runsp = null;
            //Pipeline pipeline = null;
            //try
            //{
            //   PSSnapInInfo info = runConf.AddPSSnapIn("Microsoft.Exchange.Management.PowerShell.Admin", out pssException);
            //   runsp = RunspaceFactory.CreateRunspace(runConf);
            //   System.Threading.Thread.Sleep(20000);

            //   runsp.Open();
            //   pipeline = runsp.CreatePipeline();
            //   Command command = new Command("Update-AddressList"); //powershell指令
            //   command.Parameters.Add("identity", @"\" + addresslistname2);
            //   pipeline.Commands.Add(command);
            //   pipeline.Invoke();
            //   return "人员调动成功！";
            //}
            //catch (Exception ex)
            //{
            //    Debug.WriteLine(ex.Message);
            //    return "人员调动失败！";
            //}
            //finally
            //{
            //    pipeline.Dispose();
            //    runsp.Close();
            //}

            #endregion

            try
            {
                // 生成一个连接类型，传入exchange服务器IP、将要使用的Scheme以及管理员凭据
                var connectionInfo = ExchangeScriptInit();

                using (var rs = RunspaceFactory.CreateRunspace(connectionInfo))
                {
                    var psh = PowerShell.Create();
                    psh.Runspace = rs;
                    rs.Open();
                    var sScript = "Set -Mailbox -HiddenFromAddressListsEnabled $true - Identity '" + identity + "'";
                    //Set - Mailbox - HiddenFromAddressListsEnabled $true - Identity 'test.com/xd-ad/集团公司/办公室/小伙子'
                    psh.Commands.AddScript(sScript);
                    var arr = identity.Split('/');
                    var name = arr[arr.Length - 1];
                    var addresslistname = identity.Substring(0, identity.Length - name.Length - 1);
                    var sScript1 = "Update-AddressList -identity '" + addresslistname + "'";
                    psh.Commands.AddScript(sScript1);
                    psh.Invoke();
                    rs.Close();
                    psh.Runspace.Close();
                    return true;
                }
            }
            catch (Exception)
            {
                return false;
            }
        }


        #endregion

        #region 设置用户为显示，并刷新地址列表
        //更新地址列表的操作
        public bool SetUserShow(string identity)
        {
            #region

            //RunspaceConfiguration runConf = RunspaceConfiguration.Create();
            //PSSnapInException pssException = null;
            //Runspace runsp = null;
            //Pipeline pipeline = null;
            //try
            //{
            //   PSSnapInInfo info = runConf.AddPSSnapIn("Microsoft.Exchange.Management.PowerShell.Admin", out pssException);
            //   runsp = RunspaceFactory.CreateRunspace(runConf);
            //   System.Threading.Thread.Sleep(20000);

            //   runsp.Open();
            //   pipeline = runsp.CreatePipeline();
            //   Command command = new Command("Update-AddressList"); //powershell指令
            //   command.Parameters.Add("identity", @"\" + addresslistname2);
            //   pipeline.Commands.Add(command);
            //   pipeline.Invoke();
            //   return "人员调动成功！";
            //}
            //catch (Exception ex)
            //{
            //    Debug.WriteLine(ex.Message);
            //    return "人员调动失败！";
            //}
            //finally
            //{
            //    pipeline.Dispose();
            //    runsp.Close();
            //}

            #endregion

            try
            {
                // 生成一个连接类型，传入exchange服务器IP、将要使用的Scheme以及管理员凭据
                var connectionInfo = ExchangeScriptInit();

                using (var rs = RunspaceFactory.CreateRunspace(connectionInfo))
                {
                    var psh = PowerShell.Create();
                    psh.Runspace = rs;
                    rs.Open();
                    var sScript = "Set -Mailbox -HiddenFromAddressListsEnabled $false - Identity '" + identity + "'";
                    //Set - Mailbox - HiddenFromAddressListsEnabled $true - Identity 'test.com/xd-ad/集团公司/办公室/小伙子'
                    psh.Commands.AddScript(sScript);
                    var arr = identity.Split('/');
                    var name = arr[arr.Length - 1];
                    var addresslistname = identity.Substring(0, identity.Length - name.Length - 1);
                    var sScript1 = "Update-AddressList -identity '" + addresslistname + "'";
                    psh.Commands.AddScript(sScript1);
                    psh.Invoke();
                    rs.Close();
                    psh.Runspace.Close();
                    return true;
                }
            }
            catch (Exception)
            {
                return false;
            }
        }


        #endregion 

        #region  启用或禁止显示该联系人，并更新地址列表
        //AD用户启用、禁用时，操作Exchange中的用户显示属性并更新地址列表
        // 参数addresslistname3是地址列表的名字，adName市用户的AD路径全名，如test.com/xd-ad/集团公司/办公室/spider，第三个参数表示启用为true,禁用为false
        public string UpdateAddressList3(string addresslistname3, string adName, bool enble)
        {
            #region 2007 编写方法

            #endregion

            try
            {
                // 生成一个连接类型，传入exchange服务器IP、将要使用的Scheme以及管理员凭据
                var connectionInfo = ExchangeScriptInit();

                using (var rs = RunspaceFactory.CreateRunspace(connectionInfo))
                {
                    var psh = PowerShell.Create();
                    psh.Runspace = rs;
                    rs.Open();
                    var sScript2 = "";
                    if (enble == true)
                    {
                        sScript2 = "Set-Mailbox -HiddenFromAddressListsEnabled $false -identity '" + adName + "'";

                    }
                    else
                    {
                        sScript2 = "Set-Mailbox -HiddenFromAddressListsEnabled $true -identity '" + adName + "'";
                    }
                    psh.Commands.AddScript(sScript2);

                    var sScript1 = "Update-AddressList -identity '" + addresslistname3 + "'";
                    psh.Commands.AddScript(sScript1);
                    psh.Invoke();
                    rs.Close();
                    psh.Runspace.Close();
                    return "1";
                }
            }
            catch (Exception ex)
            {
                return ex.Message;
            }
        }

        #endregion

        #region 给Exchange输入ps脚本的函数

        public WSManConnectionInfo ExchangeScriptInit()
        {
            try
            {
                var SHELL_URI = "http://schemas.microsoft.com/powershell/Microsoft.Exchange";
                // 邮件服务器地址赋值                         
                //System.Uri serverUri = new Uri(String.Format("http://Xd-svr0185.xd-ad.com.cn/PowerShell", @"xd-ad\admin"));
                //System.Uri serverUri = new Uri(String.Format("http://exchangedag.xd-ad.com.cn/PowerShell", @"xd-ad\admin"));
                //var serverUri =new Uri(string.Format("http://" + _expath + "/PowerShell", @"" + _exadmin + ""));
                var serverUri = new Uri(string.Format("http://" + _exIp + "/PowerShell"));

                PSCredential creds;
                //在内存中加密字符串
                var securePassword = new SecureString();
                foreach (var c in _expwd)
                {
                    securePassword.AppendChar(c);
                }

                //creds = new PSCredential(@"xd-ad\admin", securePassword);
                // 生成凭证（根据exchange的管理员用户名和密码）
                creds = new PSCredential(@"" + _exadmin + "", securePassword);
                // 生成一个连接类型，传入exchange服务器IP、将要使用的Scheme以及管理员凭据
                WSManConnectionInfo connectionInfo = new WSManConnectionInfo(serverUri, SHELL_URI, creds);
                return connectionInfo;
            }
            catch (Exception)
            {
                return null;
            }

            #endregion

        }
    }
}