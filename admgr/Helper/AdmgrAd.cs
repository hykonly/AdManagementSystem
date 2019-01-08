using admgr.EfModel;
using admgr.Entity;
using admgr.Models;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.DirectoryServices;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Net;
using System.Security.Principal;
using System.Text;
using System.Text.RegularExpressions;
using System.Web;
using System.Web.Mvc;

namespace ADmgr.Helper
{
    public class AdmgrAd
    {
        public DirectoryEntry _root;
        readonly string path;
        public AdmgrAd()
        {
            AdmgrModel admgrModel = new AdmgrModel();
            var admgrSystem = admgrModel.AdmgrSystems.FirstOrDefault(a => a.AdmgrSystemId == 1);
            var adDomain = "LDAP://" + admgrSystem.AdDomain;
            var adManagerUserName = admgrSystem.AdManagerUserName;
            var adManagerPwd = admgrSystem.AdManagerPwd;
            admgrModel.Dispose();
            _root = new DirectoryEntry(adDomain, adManagerUserName, adManagerPwd);
            path = adDomain;
        }

        #region 登陆验证并返回对象
        /// <summary>
        /// 根据用户名和密码验证，并返回用户的对象
        /// </summary>
        /// <param name="adusername"></param>
        /// <param name="adpwd"></param>
        /// <returns></returns>
        public DirectoryEntry CheckLogin(string adusernumber, string adpwd)
        {
            //bool result = false;
            DirectoryEntry entry = null;
            try
            {
                #region 验证DC中是否存在这个用户
             
                //userRoot是用户登录后，才能使用的root
                DirectoryEntry userRoot = new DirectoryEntry(path, adusernumber, adpwd);
                if (userRoot == null) throw new Exception("登录错误");
                DirectorySearcher searcher = new DirectorySearcher(userRoot);
                searcher.Filter = "(sAMAccountName=" + adusernumber + ")";
                entry = searcher.FindOne().GetDirectoryEntry();
                string accountexpires;
                //Bind to the native AdsObject to force authentication.
                //object obj = entry.NativeObject;

                #region 验证用户是否被禁用或者用户的过期时间已到
                if (entry != null)
                {
                    //DirectoryEntry user = result.GetDirectoryEntry();
                    //result = true;
                    // 如果账户被禁用(测试后，正常用户的val值为66048，被禁用户为66050)，那么直接设置results = false;
                    int val = (int)entry.Properties["userAccountControl"].Value;
                    if (val == 66050)
                    {
                        entry = null;
                    }
                    else if (entry.Properties.Contains("accountExpires"))
                    {
                        var asLong = LargeToLong.ConvertLargeIntegerToLong(entry.Properties["accountExpires"].Value);

                        if (asLong == long.MaxValue || asLong <= 0 || DateTime.MaxValue.ToFileTime() <= asLong)
                        {
                            accountexpires = DateTime.MaxValue.ToString(CultureInfo.CurrentCulture);
                        }
                        else
                        {
                            accountexpires = DateTime.FromFileTime(asLong).ToString(CultureInfo.CurrentCulture);
                        }
                        DateTime nowdate = DateTime.Now;
                        DateTime tempdate = Convert.ToDateTime(accountexpires);
                        if (DateTime.Compare(tempdate, nowdate) <= 0)
                        {
                            entry = null;
                        }
                    }
                }
                #endregion

                #endregion

            }
            catch (Exception ex)
            {
                entry = null;
            }
            finally
            {
            }
            return entry;
        }
        #endregion

        #region 用户的操作

        #region 查询所有公司部门下的用户
        //除了被禁目录下的用户
        public List<AdUser> GetAllAdUsers()
        {
            List<AdUser> adUserList = new List<AdUser>();
            try
            {
                DirectorySearcher deSearch = new DirectorySearcher(_root);
                deSearch.Filter = "(objectClass=organizationalUnit)";
                SearchResultCollection searchResultCollection = deSearch.FindAll();
                foreach (SearchResult companyResult in searchResultCollection)
                {
                    DirectoryEntry companyEntry = companyResult.GetDirectoryEntry();
                    var companyName = companyEntry.Properties["name"][0].ToString();
                    //如果是公司的话，distinguishedName会类似OU=0100-01000.集团公司,DC=test,DC=com这种结构，被分成的字符串的数组长度为3
                    var distinguishedNameArray = companyEntry.Properties["distinguishedName"][0].ToString().Split(',');
                    if (Regex.IsMatch(companyName, "[0-9]") && distinguishedNameArray.Length == 3 && !companyName.Contains("帐号回收站"))
                    {
                        DirectorySearcher deptSearch = new DirectorySearcher(companyEntry);
                        deptSearch.Filter = "(objectClass=organizationalUnit)";
                        SearchResultCollection deptSearchResultCollection = deptSearch.FindAll();
                        foreach (SearchResult deptResult in deptSearchResultCollection)
                        {
                            DirectoryEntry deptEntry = deptResult.GetDirectoryEntry();
                            var name = deptEntry.Properties["name"][0].ToString();
                            var deptDistinguishedNameArray = deptEntry.Properties["distinguishedName"][0].ToString().Split(',');
                            if (Regex.IsMatch(name, "[0-9]") && deptDistinguishedNameArray.Length == 4)
                            {
                                var tempAdUserList = GetAdUsers(deptEntry);
                                if (tempAdUserList != null) adUserList.AddRange(tempAdUserList);
                            }
                        }
                    }
                }
            }
            catch
            {
                return null;
            }
            return adUserList;
        }
        #endregion

        #region 新增用户对象
        /// <summary>
        /// 创建用户--部分
        /// </summary>
        /// <param name="parentEntry"></param>
        /// <param name="newAdName"></param>
        /// <param name="principalName"></param>
        /// <returns></returns>
        public DirectoryEntry CreateUser(AdUserPart AdUserPart)
        {

            DirectoryEntry userEntry = null;
            try
            {
                //根据adUser的company，dept信息找到DirectoryEntry
                DirectoryEntry comppanyEntry = GetOuDirectoryEtnryByEntryAndName(_root, AdUserPart.Company);
                DirectoryEntry deptEntry = GetOuDirectoryEtnryByEntryAndName(comppanyEntry, AdUserPart.Dept);
                #region 在父目录对象下面新增用户
                userEntry = deptEntry.Children.Add("CN=" + AdUserPart.AdName, "user");
                userEntry.CommitChanges();
                // 启用用户
                SetEnableAccount(userEntry, true);
                userEntry.CommitChanges();
                // 设置用户永远不能过期，将userAccountControl属性设置为 0x0202 (0x002 + 0x0200)。在十进制中，
                // 它是 514 (2 + 512)；若要启用账户且密码永不过期，请将 userAccountControl属性设置为 0x10200 (0x10000 + 0x0200)，
                // 十进制为66048。(启用：512，禁用：514， 密码永不过期：66048)
                userEntry.Properties["userAccountControl"].Value = 66048;
                userEntry.CommitChanges();
                // 设置用户的密码，新增用户的默认密码为123456
                string Password = "123456";
                userEntry.Invoke("SetPassword", new object[] { "" + Password + "" });
                //设置新增元素的属性
                bool result = SetAdUserInfoPart(userEntry, AdUserPart);
                userEntry.CommitChanges();
                userEntry.Close();
                #endregion

                #region 将用户移动到组里面
                //先找到对应的group对象
                DirectorySearcher groupSearcher = new DirectorySearcher(deptEntry);
                groupSearcher.Filter = "(&(name=" + deptEntry.Name.Substring(3) + ")(objectClass=group))";
                DirectoryEntry groupEntry = groupSearcher.FindOne().GetDirectoryEntry();
                //移动对象到公司的group对象里面去
                groupEntry.Properties["member"].Add(userEntry.Properties["distinguishedName"][0].ToString());
                groupEntry.CommitChanges();
                #endregion
            }
            catch (Exception e)
            {
                string ss = e.Message;
            }

            return userEntry;
        }
        #endregion

        #region 更新adUser信息--用户
        public bool SetAdUserInfo(DirectoryEntry entry, AdUser adUser)
        {
            var temp = false;
            try
            {
                if (entry != null)
                {
                    #region 编辑信息
                    SetProperty(entry, "physicalDeliveryOfficeName", adUser.Office);
                    //办公电话总机
                    SetProperty(entry, "telephoneNumber", adUser.Officephone);
                    //分机号
                    SetProperty(entry, "ipPhone", adUser.Ext);
                    //传真
                    SetProperty(entry, "facsimileTelephoneNumber", adUser.Fax);
                    //邮箱地址
                    SetProperty(entry, "mail", adUser.Email);
                    //移动电话号码
                    SetProperty(entry, "mobile", adUser.Mobile);
                    //直线电话号码
                    SetProperty(entry, "homePhone", adUser.Homephone);
                    //国家
                    SetProperty(entry, "co", adUser.Country);
                    //省/自治区
                    SetProperty(entry, "st", adUser.Province);
                    //市/县
                    SetProperty(entry, "l", adUser.City);
                    //邮编
                    SetProperty(entry, "postalCode", adUser.Zipcode);
                    //地址
                    SetProperty(entry, "streetAddress", adUser.Address);
                    //描述
                    SetProperty(entry, "description", adUser.Description);
                    #endregion
                    entry.CommitChanges();
                    entry.Close();
                    temp = true;
                }
            }
            catch (Exception e)
            {

            }
            finally
            {
            }

            return temp;
        }
        #endregion

        #region 重设AD用户的密码
        public bool ResetPwd(DirectoryEntry dey, string oldPwd, string newPwd)
        {
            bool result = false;
            try
            {
                #region 密码
                if (!String.IsNullOrEmpty(oldPwd) && !String.IsNullOrEmpty(newPwd))
                {
                    dey.Invoke("ChangePassword", new object[] { oldPwd, newPwd });
                    dey.CommitChanges();
                    dey.Close();
                    result = true;
                }
                #endregion
            }
            catch (Exception e)
            {
                //do nothing
            }
            return result;
        }

        public string ResetPwdByNameAndMobile(string name, string mobile)
        {
            try
            {
                DirectorySearcher searcher = new DirectorySearcher(_root);
                searcher.Filter = ("(&(mobile=" + mobile + ")(sAMAccountName=" + name + "))");
                DirectoryEntry userEntry = searcher.FindOne().GetDirectoryEntry();
                string newPwd = GenerateCheckCode(6);
                userEntry.Invoke("SetPassword", new object[] { "" + newPwd + "" });
                userEntry.CommitChanges();
                return SendMsg(mobile, newPwd);
            }
            catch
            {
                return "";
            }

        }
        #endregion

        #region 更新adUser信息--全部
        public bool SetAdAdUserInfoAll(DirectoryEntry dey, AdUser adUser)
        {
            var temp = false;
            try
            {
                #region 设置属性
                #region 设置公司部门属性
                SetProperty(dey, "company", adUser.Company);
                //部门
                SetProperty(dey, "department", adUser.Dept);
                //职位
                SetProperty(dey, "title", adUser.Title);
                // 姓名
                SetProperty(dey, "sn", adUser.Sn);
                // 工号
                SetProperty(dey, "givenName", adUser.Number);
                //显示名称
                SetProperty(dey, "displayName", adUser.DisplayName);
                //办公室
                SetProperty(dey, "physicalDeliveryOfficeName", adUser.Office);
                //电话总机
                SetProperty(dey, "telephoneNumber", adUser.Officephone);
                //分机号
                SetProperty(dey, "ipPhone", adUser.Ext);
                //直线电话号码
                SetProperty(dey, "homePhone", adUser.Homephone);
                //移动电话号码
                SetProperty(dey, "mobile", adUser.Mobile);
                //传真
                SetProperty(dey, "facsimileTelephoneNumber", adUser.Fax);
                ////电子邮件
                SetProperty(dey, "mail", adUser.Email );
                //国家
                SetProperty(dey, "co", adUser.Country);
                //所在省
                SetProperty(dey, "st", adUser.Province);
                //  所在市
                SetProperty(dey, "l", adUser.City);
                //邮编
                SetProperty(dey, "postalCode", adUser.Zipcode);
                // 地址
                SetProperty(dey, "streetAddress", adUser.Address);
                dey.CommitChanges();
                #endregion

                if (adUser.AccountExpires != "")
                {
                    DateTime dt = Convert.ToDateTime(adUser.AccountExpires);
                    // long longAE = dt.Ticks;
                    //SetProperty(dey, "AccountExpirationDate", dt.ToString());
                    dey.InvokeSet("AccountExpirationDate", new object[] { new DateTime(dt.Year, dt.Month, dt.Day) });
                    dey.CommitChanges();
                }
                dey.Rename("cn=" + adUser.DisplayName + "");
                dey.CommitChanges();
                SetProperty(dey, "sAMAccountName", adUser.AdName);
                dey.CommitChanges();
                #region 启/禁用用户
                if (adUser.Disble == false)
                {
                    SetEnableAccount(dey, false);
                }
                else
                {
                    SetEnableAccount(dey, true);
                }
                #endregion
                temp = true;
                #endregion
            }
            catch
            {
                //Debug.Write(ex.Message + "————————" + ex.Source + "————————" + ex.StackTrace);
            }
            finally
            {
            }
            return temp;
        }
        #endregion

        #region 更新adUser信息--部分

        public bool SetAdUserInfoPart(DirectoryEntry userEntry ,AdUserPart adUserPart)
        {
            var temp = false;
            try
            {
                #region 设置属性
                #region 设置公司部门属性
                SetProperty(userEntry, "company", adUserPart.Company);
                //部门
                SetProperty(userEntry, "department", adUserPart.Dept);
                //职位
                SetProperty(userEntry, "title", adUserPart.Title);
                // 姓名
                SetProperty(userEntry, "sn", adUserPart.Sn);
                // 工号
                SetProperty(userEntry, "givenName", adUserPart.Number);
                //显示名称
                SetProperty(userEntry, "displayName", adUserPart.DisplayName);
                //移动电话号码
                SetProperty(userEntry, "mobile", adUserPart.Mobile);
                ////电子邮件
                SetProperty(userEntry, "mail", adUserPart.Email );
                userEntry.CommitChanges();
                #endregion

                userEntry.Rename("cn=" + adUserPart.DisplayName + "");
                userEntry.CommitChanges();
                SetProperty(userEntry, "sAMAccountName", adUserPart.AdName);
                userEntry.CommitChanges();
                #region 启/禁用用户
                if (adUserPart.Disble == false)
                {
                    SetEnableAccount(userEntry, false);
                }
                else
                {
                    SetEnableAccount(userEntry, true);
                }
                #endregion
                temp = true;
                #endregion
            }
            catch
            {
                //Debug.Write(ex.Message + "————————" + ex.Source + "————————" + ex.StackTrace);
            }
            finally
            {
            }
            return temp;
        }
        #endregion

        #region 删除用户
        //这里删除操作是将用户移动到数据库中定义的DirectoryEntry里
        //顺便将该父目录下组中的用户给移除
        public bool DeleteAdUser(DirectoryEntry userEntry)
        {
            try
            {
                AdmgrModel admgrModel = new AdmgrModel();
                string ouRecycleBinName = admgrModel.AdmgrSystems.FirstOrDefault(a => a.AdmgrSystemId == 1).AccountRecycleBin;
                DirectoryEntry ouRecycleBinEntry = GetOuDirectoryEtnryByName(ouRecycleBinName);
                //将用户移动到回收AD目录
                userEntry.MoveTo(ouRecycleBinEntry);
                #region 将部门组中的用户删除
                DeleteUserFromGroup(userEntry.Parent, userEntry);
                #endregion
                SetEnableAccount(userEntry, false);
                ouRecycleBinEntry.Dispose();
                return true;
            }
            catch
            {
                return false;
            }
        }
        #endregion

        #region 获取所有职务
        public List<TitleViewModel> GetTitles()
        {
            return new List<TitleViewModel>
            {
                new TitleViewModel { TitleId = "1",TitleName = "普通员工S"},
                new TitleViewModel { TitleId = "2", TitleName = "临时员工T" },
                new TitleViewModel { TitleId = "3", TitleName = "专用邮箱Z" },
                new TitleViewModel { TitleId = "4", TitleName = "特殊账户X" },
                new TitleViewModel { TitleId = "5", TitleName = "领导A" },
                new TitleViewModel { TitleId = "6", TitleName = "领导B" },
                new TitleViewModel { TitleId = "7", TitleName = "领导C" },
                new TitleViewModel { TitleId = "8", TitleName = "领导D" },
                new TitleViewModel { TitleId = "9", TitleName = "领导E" },
                new TitleViewModel { TitleId = "10", TitleName = "领导F" },
                new TitleViewModel { TitleId = "11", TitleName = "领导G" },
                new TitleViewModel { TitleId = "12", TitleName = "领导H" },
                new TitleViewModel { TitleId = "13", TitleName = "领导I" },
                new TitleViewModel { TitleId = "14", TitleName = "领导J" },
                new TitleViewModel { TitleId = "15", TitleName = "领导K" },
                new TitleViewModel { TitleId = "16", TitleName = "领导L" },
                new TitleViewModel { TitleId = "17", TitleName = "领导M" },
                new TitleViewModel { TitleId = "18", TitleName = "领导N" },
                new TitleViewModel { TitleId = "19", TitleName = "领导O" },
                new TitleViewModel { TitleId = "20", TitleName = "领导P" },
            };
        }
        #endregion

        #region UserEntry->List<AdUser>
        /// <summary>
        /// 查询一个entry下的的所有用户
        /// </summary>
        /// <param name="entry"></param>
        /// <returns></returns>
        public List<AdUser> GetAdUsers(DirectoryEntry entry)
        {
            List<AdUser> adAdUserInfoList = new List<AdUser>();
            #region 取得用户的信息
            DirectorySearcher Dsearch = new DirectorySearcher(entry);
            Dsearch.Filter = ("(objectClass=person)");
            SearchResultCollection results = Dsearch.FindAll();
            if (results != null)
            {
                foreach (System.DirectoryServices.SearchResult resEnt in Dsearch.FindAll())
                {
                    AdUser adUserInfo = new AdUser();
                    try
                    {
                        #region 取得用户的所有信息
                        DirectoryEntry user = resEnt.GetDirectoryEntry();

                        //公司
                        if (user.Properties.Contains("company"))
                        {
                            adUserInfo.Company = user.Properties["company"][0].ToString();
                        }

                        //部门
                        if (user.Properties.Contains("department"))
                        {
                            adUserInfo.Dept = user.Properties["department"][0].ToString();
                        }

                        //职务
                        if (user.Properties.Contains("title"))
                        {
                            adUserInfo.Title = user.Properties["title"][0].ToString();
                        }
                        else
                        {
                            if (user.Properties.Contains("name"))
                            {
                                if (user.Properties["name"][0].ToString().Contains("-"))
                                {
                                    #region 职务
                                    if (user.Properties["name"][0].ToString().Substring(user.Properties["name"][0].ToString().IndexOf("-") - 1, 1).Contains("A"))
                                    {

                                        adUserInfo.Title = "领导A";
                                    }

                                    else if (user.Properties["name"][0].ToString().Substring(user.Properties["name"][0].ToString().IndexOf("-") - 1, 1).Contains("B"))
                                    {

                                        adUserInfo.Title = "领导B";
                                    }

                                    else if (user.Properties["name"][0].ToString().Substring(user.Properties["name"][0].ToString().IndexOf("-") - 1, 1).Contains("C"))
                                    {

                                        adUserInfo.Title = "领导C";
                                    }
                                    else if (user.Properties["name"][0].ToString().Substring(user.Properties["name"][0].ToString().IndexOf("-") - 1, 1).Contains("D"))
                                    {

                                        adUserInfo.Title = "领导D";
                                    }
                                    else if (user.Properties["name"][0].ToString().Substring(user.Properties["name"][0].ToString().IndexOf("-") - 1, 1).Contains("E"))
                                    {

                                        adUserInfo.Title = "领导E";
                                    }
                                    else if (user.Properties["name"][0].ToString().Substring(user.Properties["name"][0].ToString().IndexOf("-") - 1, 1).Contains("F"))
                                    {

                                        adUserInfo.Title = "领导F";
                                    }
                                    else if (user.Properties["name"][0].ToString().Substring(user.Properties["name"][0].ToString().IndexOf("-") - 1, 1).Contains("G"))
                                    {

                                        adUserInfo.Title = "领导G";
                                    }
                                    else if (user.Properties["name"][0].ToString().Substring(user.Properties["name"][0].ToString().IndexOf("-") - 1, 1).Contains("H"))
                                    {

                                        adUserInfo.Title = "领导H";
                                    }
                                    else if (user.Properties["name"][0].ToString().Substring(user.Properties["name"][0].ToString().IndexOf("-") - 1, 1).Contains("I"))
                                    {

                                        adUserInfo.Title = "领导I";
                                    }
                                    else if (user.Properties["name"][0].ToString().Substring(user.Properties["name"][0].ToString().IndexOf("-") - 1, 1).Contains("J"))
                                    {

                                        adUserInfo.Title = "领导J";
                                    }
                                    else if (user.Properties["name"][0].ToString().Substring(user.Properties["name"][0].ToString().IndexOf("-") - 1, 1).Contains("K"))
                                    {

                                        adUserInfo.Title = "领导K";
                                    }
                                    else if (user.Properties["name"][0].ToString().Substring(user.Properties["name"][0].ToString().IndexOf("-") - 1, 1).Contains("L"))
                                    {

                                        adUserInfo.Title = "领导L";
                                    }
                                    else if (user.Properties["name"][0].ToString().Substring(user.Properties["name"][0].ToString().IndexOf("-") - 1, 1).Contains("M"))
                                    {

                                        adUserInfo.Title = "领导M";
                                    }
                                    else if (user.Properties["name"][0].ToString().Substring(user.Properties["name"][0].ToString().IndexOf("-") - 1, 1).Contains("N"))
                                    {

                                        adUserInfo.Title = "领导N";
                                    }
                                    else if (user.Properties["name"][0].ToString().Substring(user.Properties["name"][0].ToString().IndexOf("-") - 1, 1).Contains("O"))
                                    {

                                        adUserInfo.Title = "领导O";
                                    }
                                    else if (user.Properties["name"][0].ToString().Substring(user.Properties["name"][0].ToString().IndexOf("-") - 1, 1).Contains("P"))
                                    {

                                        adUserInfo.Title = "领导P";
                                    }
                                    else if (user.Properties["name"][0].ToString().Substring(user.Properties["name"][0].ToString().IndexOf("-") - 1, 1).Contains("S"))
                                    {

                                        adUserInfo.Title = "普通员工S";
                                    }
                                    else if (user.Properties["name"][0].ToString().Substring(user.Properties["name"][0].ToString().IndexOf("-") - 1, 1).Contains("T"))
                                    {

                                        adUserInfo.Title = "临时员工T";
                                    }
                                    else
                                    {
                                        adUserInfo.Title = "普通员工S";
                                    }
                                }
                                #endregion
                            }
                        }

                        //姓名
                        if (user.Properties.Contains("sn"))
                        {
                            adUserInfo.Sn = user.Properties["sn"][0].ToString();
                        }
                        //工号
                        if (user.Properties.Contains("givenName"))
                        {
                            adUserInfo.Number = user.Properties["givenName"][0].ToString();
                        }
                        //域账号
                        if (user.Properties.Contains("sAMAccountName"))
                        {
                            adUserInfo.AdName = user.Properties["sAMAccountName"][0].ToString();
                        }

                        //是否禁用
                        if (user.Properties.Contains("userAccountControl"))
                        {
                            //if (Convert.ToInt32(user.Properties["userAccountControl"][0].ToString()) == 546 || Convert.ToInt32(user.Properties["userAccountControl"][0].ToString()) == 514 || Convert.ToInt32(user.Properties["userAccountControl"][0].ToString()) == 66050)
                            int num = Convert.ToInt32(user.Properties["userAccountControl"][0].ToString());
                            int[] numbers = new int[] { 8388608, 4194304, 2097152, 1048576, 524288, 262144, 131072, 65536, 8192, 4096, 2048, 512, 256, 128, 64, 32, 16, 8, 2, 1 };
                            for (int i = 0; i <= numbers.Length; i++)
                            {
                                if (numbers[i] <= num)
                                {
                                    if (numbers[i] != 2)
                                    {
                                        num -= numbers[i];

                                        if (num == 2)
                                        {
                                            adUserInfo.Disble = true;
                                            break;
                                        }
                                        else if (num == 0)
                                        {
                                            adUserInfo.Disble = false;
                                            break;
                                        }
                                    }
                                    else
                                    {
                                        adUserInfo.Disble = true;
                                        break;
                                    }
                                }
                            }
                        }

                        //过期时间
                        if (user.Properties.Contains("accountExpires"))
                        {
                            var asLong = LargeToLong.ConvertLargeIntegerToLong(user.Properties["accountExpires"].Value);

                            if (asLong == long.MaxValue || asLong <= 0 || DateTime.MaxValue.ToFileTime() <= asLong)
                            {
                                adUserInfo.AccountExpires = DateTime.MaxValue.ToString();
                            }
                            else
                            {
                                adUserInfo.AccountExpires = DateTime.FromFileTime(asLong).ToString();
                            }

                        }
                        //显示名称
                        if (user.Properties.Contains("displayName"))
                        {
                            adUserInfo.DisplayName = user.Properties["displayName"][0].ToString();
                        }

                        //办公室
                        if (user.Properties.Contains("physicalDeliveryOfficeName"))
                        {
                            adUserInfo.Office = user.Properties["physicalDeliveryOfficeName"][0].ToString();
                        }
                        //办公电话总机
                        if (user.Properties.Contains("telephoneNumber"))
                        {
                            adUserInfo.Officephone = user.Properties["telephoneNumber"][0].ToString();
                        }
                        //分机
                        if (user.Properties.Contains("ipPhone"))
                        {
                            adUserInfo.Ext = user.Properties["ipPhone"][0].ToString();
                        }

                        //直线电话号码
                        if (user.Properties.Contains("homePhone"))
                        {
                            adUserInfo.Homephone = user.Properties["homePhone"][0].ToString();
                        }
                        //移动电话号码
                        if (user.Properties.Contains("mobile"))
                        {
                            adUserInfo.Mobile = user.Properties["mobile"][0].ToString();
                        }
                        //传真
                        if (user.Properties.Contains("facsimileTelephoneNumber"))
                        {
                            adUserInfo.Fax = user.Properties["facsimileTelephoneNumber"][0].ToString();
                        }
                        //电子邮件
                        if (user.Properties.Contains("mail"))
                        {
                            adUserInfo.Email = user.Properties["mail"][0].ToString();
                        }
                        //国家
                        if (user.Properties.Contains("co"))
                        {
                            adUserInfo.Country = user.Properties["co"][0].ToString();
                        }
                        //家庭所在省
                        if (user.Properties.Contains("st"))
                        {
                            adUserInfo.Province = user.Properties["st"][0].ToString();
                        }
                        //  家庭所在市
                        if (user.Properties.Contains("l"))
                        {
                            adUserInfo.City = user.Properties["l"][0].ToString();
                        }
                        //家庭邮编
                        if (user.Properties.Contains("postalCode"))
                        {
                            adUserInfo.Zipcode = user.Properties["postalCode"][0].ToString();
                        }
                        // 地址
                        if (user.Properties.Contains("streetAddress"))
                        {
                            adUserInfo.Address = user.Properties["streetAddress"][0].ToString();
                        }
                        #endregion
                    }
                    catch (Exception ex)
                    {
                        Debug.WriteLine(ex.Message);
                        break;
                    }
                    finally
                    {
                        Dsearch.Dispose();
                        entry.Close();
                    }
                    adAdUserInfoList.Add(adUserInfo);
                }
            }
            #endregion
            return adAdUserInfoList;
        }



        /// <summary>
        /// 查询一个UsernEntry的parent的parent的所有用户
        /// </summary>
        /// <param name="userEntry"></param>
        /// <returns></returns>
        public List<AdUser> GetCompanyAdUsers(DirectoryEntry userEntry)
        {
            try
            {
                List<AdUser> adUserList = GetAdUsers(userEntry.Parent.Parent);
                return adUserList;
            }
            catch
            {
                return null;
            }

        }
        #endregion

        #region 查询过期用户
        public List<AdUser> SearchaccountExpiresr(DirectoryEntry entry)
        {
            DateTime nowdate = DateTime.Now;
            List<AdUser> adUserInfoList = GetAdUsers(entry);
            foreach (var item in adUserInfoList)
            {
                var accountexpires = item.AccountExpires;
                DateTime tempdate = Convert.ToDateTime(accountexpires);
                if (DateTime.Compare(tempdate, nowdate) > 0)
                {
                    adUserInfoList.Remove(item);
                }
            }
            return adUserInfoList;
        }
        #endregion

        #region 查询用户头像相对路径
        // 这里需要更改为AD属性中的头像
        public byte[] GetAvatarPath(DirectoryEntry userEntry)
        {
            try
            {
                byte[] imgData = null;
                if (userEntry.Properties.Contains("thumbnailPhoto") && userEntry.Properties["thumbnailPhoto"][0] != null)
                {
                    imgData = (byte[])userEntry.Properties["thumbnailPhoto"][0];
                }
                if (imgData != null)
                {
                    //var relatvePath = "/Content/avatarPath/DefaultUserPhoto.jpg";
                    //System.IO.MemoryStream ms = new System.IO.MemoryStream(imgData);
                    //System.Drawing.Image img = System.Drawing.Image.FromStream(ms);
                    //relatvePath = "/Content/photo/Avatar/" + DateTime.Now.ToString("ddHHmmssff") + ".jpg";
                    //img.Save(relatvePath);
                }
                return imgData;
            }
            catch(Exception ex)
            {
                return null;
            }

        }
        #endregion

        #region adname->DirectoryEntry
        public DirectoryEntry GetDirectoryEntryByAdName(string adname)
        {
            try
            {
                DirectorySearcher search = new DirectorySearcher(_root);
                search.Filter = "(&(sAMAccountName=" + adname + ")(objectClass=person))";
                search.PropertiesToLoad.Add("cn");
                DirectoryEntry user = search.FindOne().GetDirectoryEntry();
                return user;
            }
            catch
            {
                return null;

            }

        }
        #endregion

        #region OuName

        #endregion

        #region  displayName->AdUser
        public AdUser GetAdUserByDisplayName(string displayName)
        {
            AdUser AdUserInfo = new AdUser();
            try
            {
                DirectorySearcher search = new DirectorySearcher(_root);
                search.Filter = "(&(displayName=" + displayName + ")(objectClass=person))";
                search.PropertiesToLoad.Add("cn");
                DirectoryEntry user = search.FindOne().GetDirectoryEntry();
                #region 域帐号
                if (user.Properties.Contains("name"))
                {
                    AdUserInfo.AdName = user.Properties["name"][0].ToString();
                    //AdUserInfo.Add(user.Properties["givenName"][0].ToString());
                }
                #endregion

                #region 工号
                if (user.Properties.Contains("givenName"))
                {
                    AdUserInfo.Number = user.Properties["givenName"][0].ToString();
                    //AdUserInfo.Add(user.Properties["givenName"][0].ToString());
                }
                #endregion

                #region Last Name（在域里的字段是姓）
                if (user.Properties.Contains("sn"))
                {
                    AdUserInfo.Sn = user.Properties["sn"][0].ToString();
                    //AdUserInfo.Add(user.Properties["sn"][0].ToString());
                }
                #endregion

                #region 显示名称
                if (user.Properties.Contains("displayName"))
                {
                    AdUserInfo.DisplayName = user.Properties["displayName"][0].ToString();
                    //AdUserInfo.Add(user.Properties["displayName"][0].ToString());
                }

                #endregion

                #region 办公室地点
                if (user.Properties.Contains("physicalDeliveryOfficeName"))
                {
                    AdUserInfo.Office = user.Properties["physicalDeliveryOfficeName"][0].ToString();
                    //AdUserInfo.Add(user.Properties["physicalDeliveryOfficeName"][0].ToString());
                }
                #endregion

                #region 办公电话总机
                if (user.Properties.Contains("telephoneNumber"))
                {
                    AdUserInfo.Officephone = user.Properties["telephoneNumber"][0].ToString();
                    //AdUserInfo.Add(user.Properties["telephoneNumber"][0].ToString());
                }
                #endregion

                #region 描述
                if (user.Properties.Contains("description"))
                {
                    AdUserInfo.Description = user.Properties["description"][0].ToString();

                    //AdUserInfo.Add(user.Properties["telephoneNumber"][0].ToString());
                }
                #endregion

                #region 分机
                if (user.Properties.Contains("ipPhone"))
                {
                    AdUserInfo.Ext = user.Properties["ipPhone"][0].ToString();
                    //AdUserInfo.Add(user.Properties["ipPhone"][0].ToString());
                }
                #endregion

                #region 直线电话号码
                if (user.Properties.Contains("homePhone"))
                {
                    AdUserInfo.Homephone = user.Properties["homePhone"][0].ToString();
                    //AdUserInfo.Add(user.Properties["homePhone"][0].ToString());
                }
                #endregion

                #region 移动电话号码
                if (user.Properties.Contains("mobile"))
                {
                    AdUserInfo.Mobile = user.Properties["mobile"][0].ToString();
                    //AdUserInfo.Add(user.Properties["mobile"][0].ToString());
                }
                #endregion

                #region 传真
                if (user.Properties.Contains("facsimileTelephoneNumber"))
                {
                    AdUserInfo.Fax = user.Properties["facsimileTelephoneNumber"][0].ToString();
                    //AdUserInfo.Add(user.Properties["facsimileTelephoneNumber"][0].ToString());
                }
                #endregion

                #region 电子邮件
                if (user.Properties.Contains("mail"))
                {
                    AdUserInfo.Email = user.Properties["mail"][0].ToString();
                    //AdUserInfo.Add(user.Properties["mail"][0].ToString());
                }
                #endregion

                #region 国家
                if (user.Properties.Contains("co"))
                {
                    AdUserInfo.Country = user.Properties["co"][0].ToString();
                    //AdUserInfo.Add(user.Properties["co"][0].ToString());
                }
                #endregion

                #region 省/自治区
                if (user.Properties.Contains("st"))
                {
                    AdUserInfo.Province = user.Properties["st"][0].ToString();
                    //AdUserInfo.Add(user.Properties["st"][0].ToString());
                }
                #endregion

                #region 市/县
                if (user.Properties.Contains("l"))
                {
                    AdUserInfo.City = user.Properties["l"][0].ToString();
                    //AdUserInfo.Add(user.Properties["l"][0].ToString());
                }
                #endregion

                #region 邮编
                if (user.Properties.Contains("postalCode"))
                {
                    AdUserInfo.Zipcode = user.Properties["postalCode"][0].ToString();
                    //AdUserInfo.Add(user.Properties["postalCode"][0].ToString());
                }
                #endregion

                #region 地址
                if (user.Properties.Contains("streetAddress"))
                {
                    AdUserInfo.Address = user.Properties["streetAddress"][0].ToString();
                    //AdUserInfo.Add(user.Properties["streetAddress"][0].ToString());
                }
                #endregion

                #region 公司
                if (user.Properties.Contains("company"))
                {
                    AdUserInfo.Company = user.Properties["company"][0].ToString();
                    //AdUserInfo.Add(user.Properties["company"][0].ToString());
                }
                else
                {
                    if (user.Properties.Contains("distinguishedName"))
                    {
                        string[] userArrylist = user.Properties["distinguishedName"][0].ToString().Split(',');

                        #region 判断数组长度为7的时候
                        if (userArrylist.Length == 7)
                        {
                            AdUserInfo.Company = userArrylist[3];
                        }
                        #endregion

                        #region 判断数组长度为6的时候
                        else if (userArrylist.Length == 6)
                        {
                            AdUserInfo.Company = userArrylist[2];
                        }
                        #endregion

                        #region 判断数组长度为5的时候
                        else if (userArrylist.Length == 5)
                        {
                            AdUserInfo.Company =
                                userArrylist[1].Substring(
                                   userArrylist[1].IndexOf("=") + 1);
                        }
                        #endregion
                    }
                }
                #endregion

                #region 部门
                if (user.Properties.Contains("department"))
                {
                    AdUserInfo.Dept = user.Properties["department"][0].ToString();
                }
                else
                {
                    if (user.Properties.Contains("distinguishedName"))
                    {
                        string[] userArrylist = user.Properties["distinguishedName"][0].ToString().Split(',');
                        user.Properties["distinguishedName"][0].ToString();

                        #region 判断数组长度为7的时候
                        if (userArrylist.Length == 6)
                        {
                            AdUserInfo.Dept = userArrylist[1].Substring(userArrylist[1].IndexOf("=") + 1);

                        }
                        #endregion

                        #region 判断数组长度为6的时候
                        else if (userArrylist.Length == 5)
                        {
                            AdUserInfo.Dept = "";

                        }
                        #endregion

                    }
                    else
                    {
                        AdUserInfo.Dept = "";
                    }

                }
                #endregion

                #region 职务
                if (user.Properties.Contains("title"))
                {
                    AdUserInfo.Title = user.Properties["title"][0].ToString();
                }
                else
                {
                    if (user.Properties.Contains("displayName") && user.Properties.Contains("displayName").ToString().Contains("-"))
                    {
                        if (user.Properties["displayName"][0].ToString().Substring(user.Properties["displayName"][0].ToString().IndexOf("-") - 1, 1).Contains("A"))
                        {
                            AdUserInfo.Title = "领导A";
                        }

                        else if (user.Properties["displayName"][0].ToString().Substring(user.Properties["displayName"][0].ToString().IndexOf("-") - 1, 1).Contains("B"))
                        {

                            AdUserInfo.Title = "领导B";
                        }

                        else if (user.Properties["displayName"][0].ToString().Substring(user.Properties["displayName"][0].ToString().IndexOf("-") - 1, 1).Contains("C"))
                        {

                            AdUserInfo.Title = "领导C";
                        }
                        else if (user.Properties["displayName"][0].ToString().Substring(user.Properties["displayName"][0].ToString().IndexOf("-") - 1, 1).Contains("D"))
                        {

                            AdUserInfo.Title = "领导D";
                        }
                        else if (user.Properties["displayName"][0].ToString().Substring(user.Properties["displayName"][0].ToString().IndexOf("-") - 1, 1).Contains("E"))
                        {

                            AdUserInfo.Title = "领导E";
                        }
                        else if (user.Properties["displayName"][0].ToString().Substring(user.Properties["displayName"][0].ToString().IndexOf("-") - 1, 1).Contains("F"))
                        {

                            AdUserInfo.Title = "领导F";
                        }
                        else if (user.Properties["displayName"][0].ToString().Substring(user.Properties["displayName"][0].ToString().IndexOf("-") - 1, 1).Contains("G"))
                        {

                            AdUserInfo.Title = "领导G";
                        }
                        else if (user.Properties["displayName"][0].ToString().Substring(user.Properties["displayName"][0].ToString().IndexOf("-") - 1, 1).Contains("H"))
                        {

                            AdUserInfo.Title = "领导H";
                        }
                        else if (user.Properties["displayName"][0].ToString().Substring(user.Properties["displayName"][0].ToString().IndexOf("-") - 1, 1).Contains("I"))
                        {

                            AdUserInfo.Title = "领导I";
                        }
                        else if (user.Properties["displayName"][0].ToString().Substring(user.Properties["displayName"][0].ToString().IndexOf("-") - 1, 1).Contains("J"))
                        {

                            AdUserInfo.Title = "领导J";
                        }
                        else if (user.Properties["displayName"][0].ToString().Substring(user.Properties["displayName"][0].ToString().IndexOf("-") - 1, 1).Contains("K"))
                        {

                            AdUserInfo.Title = "领导K";
                        }
                        else if (user.Properties["displayName"][0].ToString().Substring(user.Properties["displayName"][0].ToString().IndexOf("-") - 1, 1).Contains("L"))
                        {

                            AdUserInfo.Title = "领导L";
                        }
                        else if (user.Properties["displayName"][0].ToString().Substring(user.Properties["displayName"][0].ToString().IndexOf("-") - 1, 1).Contains("M"))
                        {

                            AdUserInfo.Title = "领导M";
                        }
                        else if (user.Properties["displayName"][0].ToString().Substring(user.Properties["displayName"][0].ToString().IndexOf("-") - 1, 1).Contains("N"))
                        {

                            AdUserInfo.Title = "领导N";
                        }
                        else if (user.Properties["displayName"][0].ToString().Substring(user.Properties["displayName"][0].ToString().IndexOf("-") - 1, 1).Contains("O"))
                        {

                            AdUserInfo.Title = "领导O";
                        }
                        else if (user.Properties["displayName"][0].ToString().Substring(user.Properties["displayName"][0].ToString().IndexOf("-") - 1, 1).Contains("P"))
                        {

                            AdUserInfo.Title = "领导P";
                        }
                        else if (user.Properties["displayName"][0].ToString().Substring(user.Properties["displayName"][0].ToString().IndexOf("-") - 1, 1).Contains("S"))
                        {

                            AdUserInfo.Title = "普通员工S";
                        }
                        else if (user.Properties["displayName"][0].ToString().Substring(user.Properties["displayName"][0].ToString().IndexOf("-") - 1, 1).Contains("T"))
                        {
                            AdUserInfo.Title = "临时员工T";
                        }
                        else
                        {
                            AdUserInfo.Title = "普通员工S";
                        }
                    }
                }
                #endregion

                #region 是否启用
                if (user.Properties["userAccountControl"].Value != null)
                {
                    int num = Convert.ToInt32(user.Properties["userAccountControl"].Value);
                    //目前仅存在两种情况
                    if (num == 546)
                    {
                        //如果启用且永不过期
                        AdUserInfo.Disble = true;
                    }
                    else
                    {
                        //如果被禁用
                        AdUserInfo.Disble = false;
                    }
                }
                #endregion
            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex.Message);
                AdUserInfo = null;
            }
            finally
            {
            }
            return AdUserInfo;
        }
        #endregion

        #region  adname->AdUser
        public AdUser GetAdUserByAdName(string adname)
        {
            AdUser AdUserInfo = new AdUser();
            try
            {
                DirectorySearcher search = new DirectorySearcher(_root);
                search.Filter = "(&(sAMAccountName=" + adname + ")(objectClass=person))";
                search.PropertiesToLoad.Add("cn");
                DirectoryEntry user = search.FindOne().GetDirectoryEntry();
                #region 域帐号
                if (user.Properties.Contains("name"))
                {
                    AdUserInfo.AdName = user.Properties["name"][0].ToString();
                    //AdUserInfo.Add(user.Properties["givenName"][0].ToString());
                }
                #endregion

                #region 工号
                if (user.Properties.Contains("givenName"))
                {
                    AdUserInfo.Number = user.Properties["givenName"][0].ToString();
                    //AdUserInfo.Add(user.Properties["givenName"][0].ToString());
                }
                #endregion

                #region Last Name（在域里的字段是姓）
                if (user.Properties.Contains("sn"))
                {
                    AdUserInfo.Sn = user.Properties["sn"][0].ToString();
                    //AdUserInfo.Add(user.Properties["sn"][0].ToString());
                }
                #endregion

                #region 显示名称
                if (user.Properties.Contains("displayName"))
                {
                    AdUserInfo.DisplayName = user.Properties["displayName"][0].ToString();
                    //AdUserInfo.Add(user.Properties["displayName"][0].ToString());
                }

                #endregion

                #region 办公室地点
                if (user.Properties.Contains("physicalDeliveryOfficeName"))
                {
                    AdUserInfo.Office = user.Properties["physicalDeliveryOfficeName"][0].ToString();
                    //AdUserInfo.Add(user.Properties["physicalDeliveryOfficeName"][0].ToString());
                }
                #endregion

                #region 办公电话总机
                if (user.Properties.Contains("telephoneNumber"))
                {
                    AdUserInfo.Officephone = user.Properties["telephoneNumber"][0].ToString();
                    //AdUserInfo.Add(user.Properties["telephoneNumber"][0].ToString());
                }
                #endregion

                #region 描述
                if (user.Properties.Contains("description"))
                {
                    AdUserInfo.Description = user.Properties["description"][0].ToString();

                    //AdUserInfo.Add(user.Properties["telephoneNumber"][0].ToString());
                }
                #endregion

                #region 分机
                if (user.Properties.Contains("ipPhone"))
                {
                    AdUserInfo.Ext = user.Properties["ipPhone"][0].ToString();
                    //AdUserInfo.Add(user.Properties["ipPhone"][0].ToString());
                }
                #endregion

                #region 直线电话号码
                if (user.Properties.Contains("homePhone"))
                {
                    AdUserInfo.Homephone = user.Properties["homePhone"][0].ToString();
                    //AdUserInfo.Add(user.Properties["homePhone"][0].ToString());
                }
                #endregion

                #region 移动电话号码
                if (user.Properties.Contains("mobile"))
                {
                    AdUserInfo.Mobile = user.Properties["mobile"][0].ToString();
                    //AdUserInfo.Add(user.Properties["mobile"][0].ToString());
                }
                #endregion

                #region 传真
                if (user.Properties.Contains("facsimileTelephoneNumber"))
                {
                    AdUserInfo.Fax = user.Properties["facsimileTelephoneNumber"][0].ToString();
                    //AdUserInfo.Add(user.Properties["facsimileTelephoneNumber"][0].ToString());
                }
                #endregion

                #region 电子邮件
                if (user.Properties.Contains("mail"))
                {
                    AdUserInfo.Email = user.Properties["mail"][0].ToString();
                    //AdUserInfo.Add(user.Properties["mail"][0].ToString());
                }
                #endregion

                #region 国家
                if (user.Properties.Contains("co"))
                {
                    AdUserInfo.Country = user.Properties["co"][0].ToString();
                    //AdUserInfo.Add(user.Properties["co"][0].ToString());
                }
                #endregion

                #region 省/自治区
                if (user.Properties.Contains("st"))
                {
                    AdUserInfo.Province = user.Properties["st"][0].ToString();
                    //AdUserInfo.Add(user.Properties["st"][0].ToString());
                }
                #endregion

                #region 市/县
                if (user.Properties.Contains("l"))
                {
                    AdUserInfo.City = user.Properties["l"][0].ToString();
                    //AdUserInfo.Add(user.Properties["l"][0].ToString());
                }
                #endregion

                #region 邮编
                if (user.Properties.Contains("postalCode"))
                {
                    AdUserInfo.Zipcode = user.Properties["postalCode"][0].ToString();
                    //AdUserInfo.Add(user.Properties["postalCode"][0].ToString());
                }
                #endregion

                #region 地址
                if (user.Properties.Contains("streetAddress"))
                {
                    AdUserInfo.Address = user.Properties["streetAddress"][0].ToString();
                    //AdUserInfo.Add(user.Properties["streetAddress"][0].ToString());
                }
                #endregion

                #region 公司
                if (user.Properties.Contains("company"))
                {
                    AdUserInfo.Company = user.Properties["company"][0].ToString();
                    //AdUserInfo.Add(user.Properties["company"][0].ToString());
                }
                else
                {
                    if (user.Properties.Contains("distinguishedName"))
                    {
                        string[] userArrylist = user.Properties["distinguishedName"][0].ToString().Split(',');

                        #region 判断数组长度为7的时候
                        if (userArrylist.Length == 7)
                        {
                            AdUserInfo.Company = userArrylist[3];
                        }
                        #endregion

                        #region 判断数组长度为6的时候
                        else if (userArrylist.Length == 6)
                        {
                            AdUserInfo.Company = userArrylist[2];
                        }
                        #endregion

                        #region 判断数组长度为5的时候
                        else if (userArrylist.Length == 5)
                        {
                            AdUserInfo.Company =
                                userArrylist[1].Substring(
                                   userArrylist[1].IndexOf("=") + 1);
                        }
                        #endregion
                    }
                }
                #endregion

                #region 部门
                if (user.Properties.Contains("department"))
                {
                    AdUserInfo.Dept = user.Properties["department"][0].ToString();
                }
                else
                {
                    if (user.Properties.Contains("distinguishedName"))
                    {
                        string[] userArrylist = user.Properties["distinguishedName"][0].ToString().Split(',');
                        user.Properties["distinguishedName"][0].ToString();

                        #region 判断数组长度为7的时候
                        if (userArrylist.Length == 6)
                        {
                            AdUserInfo.Dept = userArrylist[1].Substring(userArrylist[1].IndexOf("=") + 1);

                        }
                        #endregion

                        #region 判断数组长度为6的时候
                        else if (userArrylist.Length == 5)
                        {
                            AdUserInfo.Dept = "";

                        }
                        #endregion

                    }
                    else
                    {
                        AdUserInfo.Dept = "";
                    }

                }
                #endregion

                #region 职务
                if (user.Properties.Contains("title"))
                {
                    AdUserInfo.Title = user.Properties["title"][0].ToString();
                }
                else
                {
                    if (user.Properties.Contains("displayName") && user.Properties.Contains("displayName").ToString().Contains("-"))
                    {
                        if (user.Properties["displayName"][0].ToString().Substring(user.Properties["displayName"][0].ToString().IndexOf("-") - 1, 1).Contains("A"))
                        {
                            AdUserInfo.Title = "领导A";
                        }

                        else if (user.Properties["displayName"][0].ToString().Substring(user.Properties["displayName"][0].ToString().IndexOf("-") - 1, 1).Contains("B"))
                        {

                            AdUserInfo.Title = "领导B";
                        }

                        else if (user.Properties["displayName"][0].ToString().Substring(user.Properties["displayName"][0].ToString().IndexOf("-") - 1, 1).Contains("C"))
                        {

                            AdUserInfo.Title = "领导C";
                        }
                        else if (user.Properties["displayName"][0].ToString().Substring(user.Properties["displayName"][0].ToString().IndexOf("-") - 1, 1).Contains("D"))
                        {

                            AdUserInfo.Title = "领导D";
                        }
                        else if (user.Properties["displayName"][0].ToString().Substring(user.Properties["displayName"][0].ToString().IndexOf("-") - 1, 1).Contains("E"))
                        {

                            AdUserInfo.Title = "领导E";
                        }
                        else if (user.Properties["displayName"][0].ToString().Substring(user.Properties["displayName"][0].ToString().IndexOf("-") - 1, 1).Contains("F"))
                        {

                            AdUserInfo.Title = "领导F";
                        }
                        else if (user.Properties["displayName"][0].ToString().Substring(user.Properties["displayName"][0].ToString().IndexOf("-") - 1, 1).Contains("G"))
                        {

                            AdUserInfo.Title = "领导G";
                        }
                        else if (user.Properties["displayName"][0].ToString().Substring(user.Properties["displayName"][0].ToString().IndexOf("-") - 1, 1).Contains("H"))
                        {

                            AdUserInfo.Title = "领导H";
                        }
                        else if (user.Properties["displayName"][0].ToString().Substring(user.Properties["displayName"][0].ToString().IndexOf("-") - 1, 1).Contains("I"))
                        {

                            AdUserInfo.Title = "领导I";
                        }
                        else if (user.Properties["displayName"][0].ToString().Substring(user.Properties["displayName"][0].ToString().IndexOf("-") - 1, 1).Contains("J"))
                        {

                            AdUserInfo.Title = "领导J";
                        }
                        else if (user.Properties["displayName"][0].ToString().Substring(user.Properties["displayName"][0].ToString().IndexOf("-") - 1, 1).Contains("K"))
                        {

                            AdUserInfo.Title = "领导K";
                        }
                        else if (user.Properties["displayName"][0].ToString().Substring(user.Properties["displayName"][0].ToString().IndexOf("-") - 1, 1).Contains("L"))
                        {

                            AdUserInfo.Title = "领导L";
                        }
                        else if (user.Properties["displayName"][0].ToString().Substring(user.Properties["displayName"][0].ToString().IndexOf("-") - 1, 1).Contains("M"))
                        {

                            AdUserInfo.Title = "领导M";
                        }
                        else if (user.Properties["displayName"][0].ToString().Substring(user.Properties["displayName"][0].ToString().IndexOf("-") - 1, 1).Contains("N"))
                        {

                            AdUserInfo.Title = "领导N";
                        }
                        else if (user.Properties["displayName"][0].ToString().Substring(user.Properties["displayName"][0].ToString().IndexOf("-") - 1, 1).Contains("O"))
                        {

                            AdUserInfo.Title = "领导O";
                        }
                        else if (user.Properties["displayName"][0].ToString().Substring(user.Properties["displayName"][0].ToString().IndexOf("-") - 1, 1).Contains("P"))
                        {

                            AdUserInfo.Title = "领导P";
                        }
                        else if (user.Properties["displayName"][0].ToString().Substring(user.Properties["displayName"][0].ToString().IndexOf("-") - 1, 1).Contains("S"))
                        {

                            AdUserInfo.Title = "普通员工S";
                        }
                        else if (user.Properties["displayName"][0].ToString().Substring(user.Properties["displayName"][0].ToString().IndexOf("-") - 1, 1).Contains("T"))
                        {
                            AdUserInfo.Title = "临时员工T";
                        }
                        else
                        {
                            AdUserInfo.Title = "普通员工S";
                        }
                    }
                }
                #endregion

                #region 是否启用
                if (user.Properties["userAccountControl"].Value != null)
                {
                    int num = Convert.ToInt32(user.Properties["userAccountControl"].Value);
                    //目前仅存在两种情况
                    if (num == 546)
                    {
                        //如果启用且永不过期
                        AdUserInfo.Disble = true;
                    }
                    else
                    {
                        //如果被禁用
                        AdUserInfo.Disble = false;
                    }
                }
                #endregion
            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex.Message);
                AdUserInfo = null;
            }
            finally
            {
            }
            return AdUserInfo;
        }
        #endregion

        #region  DirectoryEntry->AdUser
        public AdUser GetAdUserByDirectoryEntry(DirectoryEntry user)
        {
            AdUser AdUserInfo = new AdUser();
            try
            {
                #region 域帐号
                if (user.Properties.Contains("name"))
                {
                    AdUserInfo.AdName = user.Properties["name"][0].ToString();
                    //AdUserInfo.Add(user.Properties["givenName"][0].ToString());
                }
                #endregion

                #region 工号
                if (user.Properties.Contains("givenName"))
                {
                    AdUserInfo.Number = user.Properties["givenName"][0].ToString();
                    //AdUserInfo.Add(user.Properties["givenName"][0].ToString());
                }
                #endregion

                #region Last Name（在域里的字段是姓）
                if (user.Properties.Contains("sn"))
                {
                    AdUserInfo.Sn = user.Properties["sn"][0].ToString();
                    //AdUserInfo.Add(user.Properties["sn"][0].ToString());
                }
                #endregion

                #region 显示名称
                if (user.Properties.Contains("displayName"))
                {
                    AdUserInfo.DisplayName = user.Properties["displayName"][0].ToString();
                    //AdUserInfo.Add(user.Properties["displayName"][0].ToString());
                }

                #endregion

                #region 办公室地点
                if (user.Properties.Contains("physicalDeliveryOfficeName"))
                {
                    AdUserInfo.Office = user.Properties["physicalDeliveryOfficeName"][0].ToString();
                    //AdUserInfo.Add(user.Properties["physicalDeliveryOfficeName"][0].ToString());
                }
                #endregion

                #region 办公电话总机
                if (user.Properties.Contains("telephoneNumber"))
                {
                    AdUserInfo.Officephone = user.Properties["telephoneNumber"][0].ToString();
                    //AdUserInfo.Add(user.Properties["telephoneNumber"][0].ToString());
                }
                #endregion

                #region 分机
                if (user.Properties.Contains("ipPhone"))
                {
                    AdUserInfo.Ext = user.Properties["ipPhone"][0].ToString();
                    //AdUserInfo.Add(user.Properties["ipPhone"][0].ToString());
                }
                #endregion

                #region 直线电话号码
                if (user.Properties.Contains("homePhone"))
                {
                    AdUserInfo.Homephone = user.Properties["homePhone"][0].ToString();
                    //AdUserInfo.Add(user.Properties["homePhone"][0].ToString());
                }
                #endregion

                #region 移动电话号码
                if (user.Properties.Contains("mobile"))
                {
                    AdUserInfo.Mobile = user.Properties["mobile"][0].ToString();
                    //AdUserInfo.Add(user.Properties["mobile"][0].ToString());
                }
                #endregion

                #region 传真
                if (user.Properties.Contains("facsimileTelephoneNumber"))
                {
                    AdUserInfo.Fax = user.Properties["facsimileTelephoneNumber"][0].ToString();
                    //AdUserInfo.Add(user.Properties["facsimileTelephoneNumber"][0].ToString());
                }
                #endregion

                #region 电子邮件
                if (user.Properties.Contains("mail"))
                {
                    AdUserInfo.Email = user.Properties["mail"][0].ToString();
                    //AdUserInfo.Add(user.Properties["mail"][0].ToString());
                }
                #endregion

                #region 国家
                if (user.Properties.Contains("co"))
                {
                    AdUserInfo.Country = user.Properties["co"][0].ToString();
                    //AdUserInfo.Add(user.Properties["co"][0].ToString());
                }
                #endregion

                #region 省/自治区
                if (user.Properties.Contains("st"))
                {
                    AdUserInfo.Province = user.Properties["st"][0].ToString();
                    //AdUserInfo.Add(user.Properties["st"][0].ToString());
                }
                #endregion

                #region 市/县
                if (user.Properties.Contains("l"))
                {
                    AdUserInfo.City = user.Properties["l"][0].ToString();
                    //AdUserInfo.Add(user.Properties["l"][0].ToString());
                }
                #endregion

                #region 邮编
                if (user.Properties.Contains("postalCode"))
                {
                    AdUserInfo.Zipcode = user.Properties["postalCode"][0].ToString();
                    //AdUserInfo.Add(user.Properties["postalCode"][0].ToString());
                }
                #endregion

                #region 地址
                if (user.Properties.Contains("streetAddress"))
                {
                    AdUserInfo.Address = user.Properties["streetAddress"][0].ToString();
                    //AdUserInfo.Add(user.Properties["streetAddress"][0].ToString());
                }
                #endregion

                #region 公司
                if (user.Properties.Contains("company"))
                {
                    AdUserInfo.Company = user.Properties["company"][0].ToString();
                    //AdUserInfo.Add(user.Properties["company"][0].ToString());
                }
                else
                {
                    if (user.Properties.Contains("distinguishedName"))
                    {
                        string[] userArrylist = user.Properties["distinguishedName"][0].ToString().Split(',');

                        #region 判断数组长度为7的时候
                        if (userArrylist.Length == 7)
                        {
                            AdUserInfo.Company = userArrylist[3];
                        }
                        #endregion

                        #region 判断数组长度为6的时候
                        else if (userArrylist.Length == 6)
                        {
                            AdUserInfo.Company = userArrylist[2];
                        }
                        #endregion

                        #region 判断数组长度为5的时候
                        else if (userArrylist.Length == 5)
                        {
                            AdUserInfo.Company =
                                userArrylist[1].Substring(
                                   userArrylist[1].IndexOf("=") + 1);
                        }
                        #endregion
                    }
                }
                #endregion

                #region 部门
                if (user.Properties.Contains("department"))
                {
                    AdUserInfo.Dept = user.Properties["department"][0].ToString();
                }
                else
                {
                    if (user.Properties.Contains("distinguishedName"))
                    {
                        string[] userArrylist = user.Properties["distinguishedName"][0].ToString().Split(',');
                        user.Properties["distinguishedName"][0].ToString();

                        #region 判断数组长度为7的时候
                        if (userArrylist.Length == 6)
                        {
                            AdUserInfo.Dept = userArrylist[1].Substring(userArrylist[1].IndexOf("=") + 1);

                        }
                        #endregion

                        #region 判断数组长度为6的时候
                        else if (userArrylist.Length == 5)
                        {
                            AdUserInfo.Dept = "";

                        }
                        #endregion

                    }
                    else
                    {
                        AdUserInfo.Dept = "";
                    }

                }
                #endregion

                #region 职务
                if (user.Properties.Contains("title"))
                {
                    AdUserInfo.Title = user.Properties["title"][0].ToString();
                }
                else
                {
                    if (user.Properties.Contains("displayName") && user.Properties.Contains("displayName").ToString().Contains("-"))
                    {
                        if (user.Properties["displayName"][0].ToString().Substring(user.Properties["displayName"][0].ToString().IndexOf("-") - 1, 1).Contains("A"))
                        {
                            AdUserInfo.Title = "领导A";
                        }

                        else if (user.Properties["displayName"][0].ToString().Substring(user.Properties["displayName"][0].ToString().IndexOf("-") - 1, 1).Contains("B"))
                        {

                            AdUserInfo.Title = "领导B";
                        }

                        else if (user.Properties["displayName"][0].ToString().Substring(user.Properties["displayName"][0].ToString().IndexOf("-") - 1, 1).Contains("C"))
                        {

                            AdUserInfo.Title = "领导C";
                        }
                        else if (user.Properties["displayName"][0].ToString().Substring(user.Properties["displayName"][0].ToString().IndexOf("-") - 1, 1).Contains("D"))
                        {

                            AdUserInfo.Title = "领导D";
                        }
                        else if (user.Properties["displayName"][0].ToString().Substring(user.Properties["displayName"][0].ToString().IndexOf("-") - 1, 1).Contains("E"))
                        {

                            AdUserInfo.Title = "领导E";
                        }
                        else if (user.Properties["displayName"][0].ToString().Substring(user.Properties["displayName"][0].ToString().IndexOf("-") - 1, 1).Contains("F"))
                        {

                            AdUserInfo.Title = "领导F";
                        }
                        else if (user.Properties["displayName"][0].ToString().Substring(user.Properties["displayName"][0].ToString().IndexOf("-") - 1, 1).Contains("G"))
                        {

                            AdUserInfo.Title = "领导G";
                        }
                        else if (user.Properties["displayName"][0].ToString().Substring(user.Properties["displayName"][0].ToString().IndexOf("-") - 1, 1).Contains("H"))
                        {

                            AdUserInfo.Title = "领导H";
                        }
                        else if (user.Properties["displayName"][0].ToString().Substring(user.Properties["displayName"][0].ToString().IndexOf("-") - 1, 1).Contains("I"))
                        {

                            AdUserInfo.Title = "领导I";
                        }
                        else if (user.Properties["displayName"][0].ToString().Substring(user.Properties["displayName"][0].ToString().IndexOf("-") - 1, 1).Contains("J"))
                        {

                            AdUserInfo.Title = "领导J";
                        }
                        else if (user.Properties["displayName"][0].ToString().Substring(user.Properties["displayName"][0].ToString().IndexOf("-") - 1, 1).Contains("K"))
                        {

                            AdUserInfo.Title = "领导K";
                        }
                        else if (user.Properties["displayName"][0].ToString().Substring(user.Properties["displayName"][0].ToString().IndexOf("-") - 1, 1).Contains("L"))
                        {

                            AdUserInfo.Title = "领导L";
                        }
                        else if (user.Properties["displayName"][0].ToString().Substring(user.Properties["displayName"][0].ToString().IndexOf("-") - 1, 1).Contains("M"))
                        {

                            AdUserInfo.Title = "领导M";
                        }
                        else if (user.Properties["displayName"][0].ToString().Substring(user.Properties["displayName"][0].ToString().IndexOf("-") - 1, 1).Contains("N"))
                        {

                            AdUserInfo.Title = "领导N";
                        }
                        else if (user.Properties["displayName"][0].ToString().Substring(user.Properties["displayName"][0].ToString().IndexOf("-") - 1, 1).Contains("O"))
                        {

                            AdUserInfo.Title = "领导O";
                        }
                        else if (user.Properties["displayName"][0].ToString().Substring(user.Properties["displayName"][0].ToString().IndexOf("-") - 1, 1).Contains("P"))
                        {

                            AdUserInfo.Title = "领导P";
                        }
                        else if (user.Properties["displayName"][0].ToString().Substring(user.Properties["displayName"][0].ToString().IndexOf("-") - 1, 1).Contains("S"))
                        {

                            AdUserInfo.Title = "普通员工S";
                        }
                        else if (user.Properties["displayName"][0].ToString().Substring(user.Properties["displayName"][0].ToString().IndexOf("-") - 1, 1).Contains("T"))
                        {
                            AdUserInfo.Title = "临时员工T";
                        }
                        else
                        {
                            AdUserInfo.Title = "普通员工S";
                        }
                    }
                }
                #endregion

                #region 是否启用
                if (user.Properties["userAccountControl"].Value != null)
                {
                    int num = Convert.ToInt32(user.Properties["userAccountControl"].Value);
                    //目前仅存在两种情况
                    if (num == 546)
                    {
                        //如果启用且永不过期
                        AdUserInfo.Disble = true;
                    }
                    else
                    {
                        //如果被禁用
                        AdUserInfo.Disble = false;
                    }
                }
                #endregion
            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex.Message);
                AdUserInfo = null;
            }
            finally
            {
            }
            return AdUserInfo;
        }
        #endregion

        #region 移动用户，组-->组
        public bool MoveUserFromDeptToDept(DirectoryEntry userEntry , string oldDeptName,string newDeptName)
        {
            DirectoryEntry oldDeptEntry = GetOuDirectoryEtnryByName(oldDeptName);
            DirectoryEntry newDeptEntry = GetOuDirectoryEtnryByName(newDeptName);
            DirectoryEntry oldDeptGroupEntry = GetGroupDirectoryEntryByNameAndEntry(oldDeptEntry, oldDeptEntry.Properties["name"][0].ToString());
            DirectoryEntry newDeptGroupEntry = GetGroupDirectoryEntryByNameAndEntry(newDeptEntry, newDeptEntry.Properties["name"][0].ToString());
            userEntry.MoveTo(newDeptEntry);
            userEntry.CommitChanges();
            //组的操作
            return MoveUserFromGroupToGroup(userEntry, oldDeptGroupEntry, newDeptGroupEntry);
        }
        #endregion

        #endregion

        #region 组的操作

        #region 组--添加用户
        public bool AddUserToGroup(DirectoryEntry parentEntry, DirectoryEntry userEntry)
        {
            string groupName = parentEntry.Properties["CN"].ToString();
            DirectorySearcher deSearch = new DirectorySearcher(parentEntry);
            deSearch.Filter = "(&(objectClass=group) (cn=" + groupName + "))";
            DirectoryEntry groupEntry = deSearch.FindOne().GetDirectoryEntry();
            try
            {
                //移动对象到公司的group对象里面去
                groupEntry.Properties["member"].Add(userEntry.Properties["distinguishedName"][0].ToString());
                groupEntry.CommitChanges();
                groupEntry.Close();
                return true;
            }
            catch (Exception e)
            {
                return false;
            }
        }
        #endregion

        #region 组--删除用户
        public bool DeleteUserFromGroup(DirectoryEntry parentEntry, DirectoryEntry userEntry)
        {
            try
            {
                string groupName = parentEntry.Properties["CN"].ToString();
                DirectorySearcher deSearch = new DirectorySearcher(parentEntry);
                deSearch.Filter = "(&(objectClass=group) (cn=" + groupName + "))";
                DirectoryEntry groupEntry = deSearch.FindOne().GetDirectoryEntry();
                //从公司的group对象里面移除对象
                groupEntry.Properties["member"].Remove(userEntry.Properties["distinguishedName"][0].ToString());
                groupEntry.CommitChanges();
                groupEntry.Close();
                return true;
            }
            catch (Exception e)
            {
                return false;
            }
        }
        #endregion

        #region 组--移动用户
        public bool MoveUserFromGroupToGroup(DirectoryEntry deUser,  DirectoryEntry oldGroup, DirectoryEntry newGroup)
        {
            bool isGroupMember = false;
            if (oldGroup.Name.Length > 0 && newGroup.Name.Length > 0)
            {
                try
                {
                    oldGroup.Properties["member"].Remove(deUser.Properties["distinguishedName"][0].ToString());
                    oldGroup.CommitChanges();
                    //移动对象到公司的group对象里面去
                    newGroup.Properties["member"].Add(deUser.Properties["distinguishedName"][0].ToString());
                    newGroup.CommitChanges();
                    isGroupMember = true;
                }
                catch (Exception e)
                {
                    isGroupMember = false;
                }
            }
            return isGroupMember;
        }

        #endregion

        #endregion

        #region OU的操作

        #region 公司的操作
        /// <summary>
        /// 公司的名称中要包含0-9的数字，排除被禁目录的名称
        /// </summary>
        /// <returns></returns>
        public List<CompanyViewModel> GetCompanies()
        {
            List<CompanyViewModel> models = new List<CompanyViewModel>();
            try
            {
                DirectorySearcher deSearch = new DirectorySearcher(_root);
                deSearch.Filter = "(objectClass=organizationalUnit)";
                SearchResultCollection searchResultCollection = deSearch.FindAll();
                int i = 1;
                foreach (SearchResult result in searchResultCollection)
                {
                    DirectoryEntry entry = result.GetDirectoryEntry();
                    var name = entry.Properties["name"][0].ToString();
                    //如果是公司的话，distinguishedName会类似OU=0100-01000.集团公司,DC=test,DC=com这种结构，被分成的字符串的数组长度为3
                    var distinguishedNameArray = entry.Properties["distinguishedName"][0].ToString().Split(',');
                    if (Regex.IsMatch(name, "[0-9]") && distinguishedNameArray.Length == 3 && !name.Contains("帐号回收站"))
                    {
                        CompanyViewModel model = new CompanyViewModel()
                        {
                            CompanyId = i.ToString(),
                            CompanyName = entry.Properties["name"][0].ToString()
                        };
                        i++;
                        models.Add(model);
                    }
                }
            }
            catch
            {
                return null;
            }
            return models;
        }

        public CompanyAndDeptViewModel GetCompaniesAndDepts()
        {
            CompanyAndDeptViewModel companyAndDeptViewModel = new CompanyAndDeptViewModel();
            try
            {
                List<CompanyViewModel> companyViewModelList = new List<CompanyViewModel>();
                List<DeptViewModel> deptViewModelList = new List<DeptViewModel>();
                DirectorySearcher companySearch = new DirectorySearcher(_root);
                companySearch.Filter = "(objectClass=organizationalUnit)";
                SearchResultCollection searchResultCollection = companySearch.FindAll();
                int i = 1;
                int j = 1;
                foreach (SearchResult companyResult in searchResultCollection)
                {
                    DirectoryEntry companyEntry = companyResult.GetDirectoryEntry();
                    var companyName = companyEntry.Properties["name"][0].ToString();
                    //如果是公司的话，distinguishedName会类似OU=0100-01000.集团公司,DC=test,DC=com这种结构，被分成的字符串的数组长度为3
                    var companyDistinguishedNameArray = companyEntry.Properties["distinguishedName"][0].ToString().Split(',');
                    if (Regex.IsMatch(companyName, "[0-9]") && companyDistinguishedNameArray.Length == 3 )
                    {
                        CompanyViewModel companyViewModel = new CompanyViewModel()
                        {
                            CompanyId = i.ToString(),
                            CompanyName = companyName
                        };
                        companyViewModelList.Add(companyViewModel);
                        #region 循环遍历公司中的部门
                        DirectorySearcher deptSearch = new DirectorySearcher(companyEntry);
                        deptSearch.Filter = "(objectClass=organizationalUnit)";
                        SearchResultCollection deptSearchResultCollection = deptSearch.FindAll();
                        foreach (SearchResult deptResult in deptSearchResultCollection)
                        {
                            DirectoryEntry deptEntry = deptResult.GetDirectoryEntry();
                            var deptName = deptEntry.Properties["name"][0].ToString();
                            //如果是公司的话，distinguishedName会类似OU=0100-01000.集团公司,DC=test,DC=com这种结构，被分成的字符串的数组长度为3
                            var deptDistinguishedNameArray = deptEntry.Properties["distinguishedName"][0].ToString().Split(',');
                            if (Regex.IsMatch(deptName, "[0-9]") && deptDistinguishedNameArray.Length == 4)
                            {
                                DeptViewModel deptViewModel = new DeptViewModel()
                                {
                                    CompanyId = i.ToString(),
                                    DeptId = j.ToString(),
                                    DeptName = deptName
                                };
                                j++;
                                deptViewModelList.Add(deptViewModel);
                            }
                        }
                        i++;
                        #endregion
                    }
                }
                companyAndDeptViewModel.CompanyViewModelList = companyViewModelList;
                companyAndDeptViewModel.DeptViewModelList = deptViewModelList;

            }
            catch(Exception ex)
            {
                _root.Dispose();
                return null;
            }
            return companyAndDeptViewModel;
        }

        public CompanyAndDeptViewModel GetCompaniesAndDeptsByUserEntry(DirectoryEntry userEntry)
        {
            CompanyAndDeptViewModel companyAndDeptViewModel = new CompanyAndDeptViewModel();
            try
            {
                List<CompanyViewModel> companyViewModelList = new List<CompanyViewModel>();
                List<DeptViewModel> deptViewModelList = new List<DeptViewModel>();
                int i = 1;
                int j = 1;
                DirectoryEntry companyEntry = userEntry.Parent.Parent;
                var companyName = companyEntry.Properties["name"][0].ToString();
                //如果是公司的话，distinguishedName会类似OU=0100-01000.集团公司,DC=test,DC=com这种结构，被分成的字符串的数组长度为3
                var companyDistinguishedNameArray = companyEntry.Properties["distinguishedName"][0].ToString().Split(',');
                if (Regex.IsMatch(companyName, "[0-9]") && companyDistinguishedNameArray.Length == 3)
                {
                    CompanyViewModel companyViewModel = new CompanyViewModel()
                    {
                        CompanyId = i.ToString(),
                        CompanyName = companyName
                    };
                    companyViewModelList.Add(companyViewModel);
                    #region 循环遍历公司中的部门
                    DirectorySearcher deptSearch = new DirectorySearcher(companyEntry);
                    deptSearch.Filter = "(objectClass=organizationalUnit)";
                    SearchResultCollection deptSearchResultCollection = deptSearch.FindAll();
                    foreach (SearchResult deptResult in deptSearchResultCollection)
                    {
                        DirectoryEntry deptEntry = deptResult.GetDirectoryEntry();
                        var deptName = deptEntry.Properties["name"][0].ToString();
                        //如果是公司的话，distinguishedName会类似OU=0100-01000.集团公司,DC=test,DC=com这种结构，被分成的字符串的数组长度为3
                        var deptDistinguishedNameArray = deptEntry.Properties["distinguishedName"][0].ToString().Split(',');
                        if (Regex.IsMatch(deptName, "[0-9]") && deptDistinguishedNameArray.Length == 4)
                        {
                            DeptViewModel deptViewModel = new DeptViewModel()
                            {
                                CompanyId = i.ToString(),
                                DeptId = j.ToString(),
                                DeptName = deptName
                            };
                            j++;
                            deptViewModelList.Add(deptViewModel);
                        }
                    }
                    i++;
                    #endregion
                }
                companyAndDeptViewModel.CompanyViewModelList = companyViewModelList;
                companyAndDeptViewModel.DeptViewModelList = deptViewModelList;

            }
            catch (Exception ex)
            {
                _root.Dispose();
                return null;
            }
            return companyAndDeptViewModel;
        }
        public bool AddCompany(string organizationUnitName)
        {
            try
            {
                // 新增公司OU
                DirectoryEntry companyEntry = _root.Children.Add("OU="+organizationUnitName, "organizationalUnit");
                _root.CommitChanges();
                //将新增的OU设置为防止意外删除
                bool bt = UnexpectedDeleteOu(companyEntry, true);
                // 新增公司组--group
                DirectoryEntry group = companyEntry.Children.Add("CN=" + organizationUnitName, "group");
                companyEntry.CommitChanges();
                //修改组的作用域属性
                SetProperty(group, "groupType", "-2147483640");
                group.CommitChanges();
                SetProperty(group, "sAMAccountName", organizationUnitName);
                group.CommitChanges();
                //Exchange相关操作
                //string temp = ModifyPathToIdentity(newSubEntry.Path);
                //bool r1 = adMgrExchange.AddDistributionGroupAndAddAddressList_company(temp);
                companyEntry.Dispose();
                group.Dispose();
            }
            catch(Exception ex)
            {
                return false;
            }
            return true;
        }
        /// <summary>
        /// rename company and it's group
        /// </summary>
        /// <param name="companyOldName"></param>
        /// <param name="companyNewName"></param>
        /// <returns></returns>
        public bool EditCompany(string companyOldName, string companyNewName)
        {
            try
            {
                DirectoryEntry companyEntry = GetOuDirectoryEtnryByName(companyOldName);
                companyEntry.Rename("OU="+companyNewName);
                companyEntry.CommitChanges();
                //重命名相应的组
                DirectorySearcher search2 = new DirectorySearcher(companyEntry);
                search2.Filter = "(&(name=" + companyOldName + ")(objectClass=group))";
                DirectoryEntry groupEntry = search2.FindOne().GetDirectoryEntry();
                groupEntry.Rename("CN=" + companyNewName);
                groupEntry.CommitChanges();


                companyEntry.Dispose();

                return true;
            }
            catch(Exception ex)
            {
                return false;
            }
        }
        public bool DeleteCompany(string companyName)
        {
            try
            {
                DirectoryEntry companyEntry = GetOuDirectoryEtnryByName(companyName);
                companyEntry.DeleteTree();
                companyEntry.CommitChanges();
                companyEntry.Dispose();
                return true;
            }
            catch(Exception ex)
            {
                return false;
            }
        }
        #endregion

        #region 部门的操作
        //取得所有部门
        public List<DeptViewModel> GetDeptsByCompany(CompanyViewModel companyViewModel)
        {
            var companyId = companyViewModel.CompanyId;
            var companyName = companyViewModel.CompanyName;
            List<DeptViewModel> deptViewModelList = new List<DeptViewModel>();
            DirectoryEntry companyEntry = GetOuDirectoryEtnryByName(companyName);
            #region 循环遍历公司中的部门
            DirectorySearcher deptSearch = new DirectorySearcher(companyEntry);
            deptSearch.Filter = "(objectClass=organizationalUnit)";
            SearchResultCollection deptSearchResultCollection = deptSearch.FindAll();
            int j = 1;
            foreach (SearchResult deptResult in deptSearchResultCollection)
            {
                DirectoryEntry deptEntry = deptResult.GetDirectoryEntry();
                var deptName = deptEntry.Properties["name"][0].ToString();
                //如果是公司的话，distinguishedName会类似OU=0100-01000.集团公司,DC=test,DC=com这种结构，被分成的字符串的数组长度为3
                var deptDistinguishedNameArray = deptEntry.Properties["distinguishedName"][0].ToString().Split(',');
                if (Regex.IsMatch(deptName, "[0-9]") && deptDistinguishedNameArray.Length == 4)
                {
                    DeptViewModel deptViewModel = new DeptViewModel()
                    {
                        CompanyId = companyId,
                        DeptId = j.ToString(),
                        DeptName = deptName
                    };
                    j++;
                    deptViewModelList.Add(deptViewModel);
                }
            }
            #endregion
            return deptViewModelList;
        }
        public bool AddDept(string companyName, string deptName)
        {
            try
            {
                DirectoryEntry companyEntry = GetOuDirectoryEtnryByName(companyName);
                // 新增部门OU
                DirectoryEntry deptEntry = companyEntry.Children.Add(deptName, "organizationalUnit");
                deptEntry.CommitChanges();
                //将新增的OU设置为防止意外删除
                bool bt = UnexpectedDeleteOu(companyEntry, true);
                // 新增部门组--group
                DirectoryEntry deptGroup = companyEntry.Children.Add("CN=" + deptName.Substring(3), "group");
                deptGroup.CommitChanges();
                //修改组的作用域属性
                SetProperty(deptGroup, "groupType", "-2147483640");
                deptGroup.CommitChanges();
                SetProperty(deptGroup, "sAMAccountName", deptName.Substring(deptName.IndexOf("-") + 1));
                #region 移动dept的group到company的group中
                //先找到对应的group对象
                DirectorySearcher group_searcher = new DirectorySearcher(companyEntry);
                group_searcher.Filter = "(&(name=" + companyEntry.Name.Substring(3) + ")(objectClass=group))";
                DirectoryEntry compaynGroupEntry = group_searcher.FindOne().GetDirectoryEntry();
                //移动对象到公司的group对象里面去
                compaynGroupEntry.Properties["member"].Add(deptGroup.Properties["distinguishedName"][0].ToString());
                compaynGroupEntry.CommitChanges();
                #endregion
                //Exchange相关操作
                //string temp = ModifyPathToIdentity(newSubEntry.Path);
                //bool r1 = adMgrExchange.AddDistributionGroupAndAddAddressList_company(temp);
                return true;
            }
            catch(Exception ex)
            {
                return false;
            }
        }
        /// <summary>
        /// 修改部门-1.修改部门本身，2.修改部门下的组，3.修改部门下的所有用户的显示名和部门属性
        /// </summary>
        /// <param name="companyName"></param>
        /// <param name="deptOldName"></param>
        /// <param name="deptNewName"></param>
        /// <returns></returns>
        public bool EditDept(string companyName, string deptOldName, string deptNewName)
        {
            try
            {
                DirectoryEntry companyEntry = GetOuDirectoryEtnryByName(companyName);
                DirectoryEntry deptEntry = GetOuDirectoryEtnryByEntryAndName(companyEntry, deptOldName);
                deptEntry.Rename("OU=" + deptNewName);
                deptEntry.CommitChanges();
                //重命名相应的组
                DirectorySearcher search2 = new DirectorySearcher(deptEntry);
                search2.Filter = "(&(name=" + deptOldName + ")(objectClass=group))";
                DirectoryEntry groupEntry = search2.FindOne().GetDirectoryEntry();
                groupEntry.Rename("CN=" + deptNewName);
                groupEntry.CommitChanges();
                groupEntry.Dispose();
                companyEntry.Dispose();
                // 再重命名该部门OU下所有的子元素的AD帐号中的的部门的名称
                //新的部门的名称newName
                //oldAdName找到帐号名称中'-'符号前一位到最后一位的字符串
                //拼起来就是每一帐号新的名称
                //调用重新给AD帐号重命名的方法RenameByOldAndNew，用此方法循环给OU下的所有帐号重命名
                if (deptEntry.Children != null)
                {
                    int j = 0;
                    int m = 0;
                    foreach (DirectoryEntry personEntry in deptEntry.Children)
                    {
                        if (personEntry.Name.Contains(".") && personEntry.Name.Contains("-"))
                        {
                            string oldAdName = personEntry.Name.Substring(personEntry.Name.LastIndexOf('-') - 1);
                            string newAdName = deptEntry.Name + oldAdName;
                            var temp = RenameByCn(personEntry, personEntry.Name.Substring(3), newAdName.Substring(3));
                            if (temp) j++;
                            m++;
                        }
                    }
                    if (j < m)
                    {
                        return true;
                    }
                    return false;
                }
                return true;
            }
            catch(Exception ex)
            {
                return false;
            }
        }
        public bool DeleteDept(string companyName, string deptName)
        {
            try
            {
                DirectoryEntry companyEntry = GetOuDirectoryEtnryByName(companyName);
                DirectoryEntry deptEntry = GetOuDirectoryEtnryByEntryAndName(companyEntry, deptName);
                deptEntry.DeleteTree();
                companyEntry.DeleteTree();
                deptEntry.Dispose();
                return true;
            }
            catch
            {
                return false;
            }
        }
        #endregion

        #region OuName->DirectoryEntry
        public DirectoryEntry GetOuDirectoryEtnryByName(string ouName)
        {
            try
            {
                DirectorySearcher deSearch = new DirectorySearcher(_root);
                deSearch.Filter = "(&(objectClass=organizationalUnit)(name=" + ouName + "))";
                DirectoryEntry ouEntry = deSearch.FindOne().GetDirectoryEntry();
                return ouEntry;
            }
            catch
            {
                return null;
            }

        }
        #endregion

        #region GroupName->DirectoryEntry
        public DirectoryEntry GetGroupDirectoryEntryByNameAndEntry(DirectoryEntry ouEntry,string groupName)
        {
            try
            {
                DirectorySearcher deSearch = new DirectorySearcher(ouEntry);
                deSearch.Filter = "(&(objectClass=group)(name=" + groupName + "))";
                DirectoryEntry groupEntry = deSearch.FindOne().GetDirectoryEntry();
                return groupEntry;
            }
            catch
            {
                return null;
            }

        }
        #endregion

        #region DirectoryEntry,OuName->DirectoryEntry
        public DirectoryEntry GetOuDirectoryEtnryByEntryAndName(DirectoryEntry root, string ouName)
        {
            try
            {
                DirectorySearcher deSearch = new DirectorySearcher(root);
                deSearch.Filter = "(&(objectClass=organizationalUnit)(name=" + ouName + "))";
                DirectoryEntry ouEntry = deSearch.FindOne().GetDirectoryEntry();
                return ouEntry;
            }
            catch
            {
                return null;
            }

        }
        #endregion

        #region DirectoryEntry重命名--用户
        public bool RenameByCn(DirectoryEntry user, string old_name, string new_name)
        {
            var result = false;
            try
            {
                SetProperty(user, "displayName", new_name);
                user.CommitChanges();
                SetProperty(user, "userPrincipalName", new_name );
                user.CommitChanges();
                SetProperty(user, "sAMAccountName", new_name.Substring(new_name.IndexOf('-') + 1));
                user.CommitChanges();
                user.Rename("CN=" + new_name);
                user.CommitChanges();
                result = true;

                //进行exchange相关操作
                //AdMgrExchange adMgrExchange = new AdMgrExchange(adAdminUserName, path, pwd, exadmin, expath, expwd);
                //string identity = ModifyPathToIdentity(user.Path);
                //string addresslistname = identity.Substring(0, identity.LastIndexOf('/'));
                //bool test = adMgrExchange.UpdateAddressListAfterOperate(addresslistname);
            }
            catch (Exception e)
            {
                Debug.Write(e.Message);
            }
            return result;
        }

        #endregion

        #region DirectoryEntry重命名--部门
        // 根据输入的entry和新旧名称来实现重命名，返回是否成功的结果
        // 重命名OU的同时，要重命名下面的组
        public int RenameOuDept(DirectoryEntry entry, string oldName, string newName)
        {
            int i = 0;
            DirectorySearcher search = new DirectorySearcher(entry);
            search.Filter = "(name=" + oldName.Substring(3) + ")";
            try
            {
                DirectoryEntry user = search.FindOne().GetDirectoryEntry();
                user.Rename(newName);
                user.CommitChanges();
                // 组重命名完成后，再重命名相对应的组
                DirectorySearcher search2 = new DirectorySearcher(user);
                search2.Filter = "(name=" + oldName.Substring(3) + ")";
                DirectoryEntry user2 = search2.FindOne().GetDirectoryEntry();
                user2.Rename("CN=" + newName.Substring(3));
                user2.CommitChanges();
                // 再重命名该部门OU下所有的子元素的AD帐号中的的部门的名称
                //新的部门的名称newName
                //oldAdName找到帐号名称中'-'符号前一位到最后一位的字符串
                //拼起来就是每一帐号新的名称
                //调用重新给AD帐号重命名的方法RenameByOldAndNew，用此方法循环给OU下的所有帐号重命名
                if (user.Children != null)
                {
                    int j = 0;
                    int m = 0;
                    foreach (DirectoryEntry personEntry in user.Children)
                    {
                        if (personEntry.Name.Contains(".") && personEntry.Name.Contains("-"))
                        {
                            string oldAdName = personEntry.Name.Substring(personEntry.Name.LastIndexOf('-') - 1);
                            string newAdName = user.Name + oldAdName;
                            var temp = RenameByCn(personEntry, personEntry.Name.Substring(3), newAdName.Substring(3));
                            if (temp) j++;
                            m++;
                        }
                    }
                    if (j < m)
                    {
                        i = 1;
                    }
                }
                #region Exchange操作
                //将Exchange中的元素进行修改操作
                //重命名OU完成后，要更新Exchagne地址列表中相应的OU元素，便于下面在地址列表中
                //构建一个old的ouAddress
                //AdMgrExchange adMgrExchange = new AdMgrExchange(adAdminUserName, path, pwd, exadmin, expath, expwd);
                //string identity = ModifyPathToIdentity(user.Path);
                //string oldIdentity = identity.Substring(0, identity.LastIndexOf("/") + 1) + oldName.Substring(3);
                //bool result = adMgrExchange.ModifyDistributionGroupAndModifyAddressList_dept(identity, oldIdentity);
                #endregion
            }
            catch (Exception e)
            {
                Debug.Write(e.Message);
            }

            return i;
        }

        #endregion

        #region DirectoryEntry重命名--公司
        // 根据输入的entry和新旧名称来实现重命名，返回是否成功的结果
        // 重命名OU的同时，要重命名下面的组
        public bool RenameOuCompany(DirectoryEntry entry, string old_name, string new_name)
        {

            var result = false;
            DirectorySearcher search = new DirectorySearcher(entry);
            search.Filter = "(name=" + old_name.Substring(3) + ")";
            try
            {
                DirectoryEntry user = search.FindOne().GetDirectoryEntry();
                user.Rename(new_name);
                user.CommitChanges();
                // 组重命名完成后，再重命名相对应的组
                DirectorySearcher search2 = new DirectorySearcher(user);
                search2.Filter = "(&(name=" + old_name.Substring(3) + ")(objectClass=group))";
                DirectoryEntry user2 = search2.FindOne().GetDirectoryEntry();
                user2.Rename("CN=" + new_name.Substring(3));
                user2.CommitChanges();
                // 进行Exchange的相关操作
                //AdMgrExchange adMgrExchange = new AdMgrExchange(adAdminUserName, path, pwd, exadmin, expath, expwd);
                //bool result = adMgrExchange.ModifyDistributionGroupAndModifyAddressList_company(user2.Path);
                result = true;
            }
            catch (Exception e)
            {
                Debug.Write(e.Message);
            }
            return result;
        }

        #endregion

        #endregion

        #region 其他辅助函数


        #region 将DirectoryEntry的path属性更改为identity需要的字符串
        public string ModifyPathToIdentity(string path)
        {
            string identity = String.Empty;
            var arr = path.Split(',');
            for (int i = 0; i < arr.Length; i++)
            {
                var arr2 = arr[i].Split('=');
                arr[i] = arr2[1];
            }
            StringBuilder temp = new StringBuilder();
            string temp2 = arr[arr.Length - 2] + "." + arr[arr.Length - 1] + "/";
            temp.Append(temp2);
            for (int i = arr.Length - 3; i >= 0; i--)
            {
                temp.Append(arr[i] + "/");
            }
            string result = temp.ToString().Substring(0, temp.ToString().Length - 1);
            return result;
        }
        #endregion

        #region AD操作--设置属性
        public void SetProperty(DirectoryEntry de, string PropertyName, string PropertyValue)
        {
            if (PropertyValue != null)
            {
                if (de.Properties.Contains(PropertyName))
                {
                    if (PropertyValue != "")
                    {
                        de.Properties[PropertyName][0] = PropertyValue;
                    }
                    else
                    {
                        de.Properties[PropertyName].Clear();
                    }
                }
                else
                {
                    if (PropertyValue != "")
                    {
                        de.Properties[PropertyName].Add(PropertyValue);
                    }
                }
            }
        }
        #endregion

        #region 设置ou属性--防止意外删除
        // 默认返回true ，防止意外删除
        public bool UnexpectedDeleteOu(DirectoryEntry directoryEntry, bool preset)
        {
            bool result = true;
            try
            {
                System.Security.Principal.IdentityReference newOwner = new System.Security.Principal.NTAccount("Everyone").Translate(typeof(System.Security.Principal.SecurityIdentifier));
                ActiveDirectoryAccessRule rule = new ActiveDirectoryAccessRule(newOwner, ActiveDirectoryRights.Delete | ActiveDirectoryRights.DeleteChild | ActiveDirectoryRights.DeleteTree, System.Security.AccessControl.AccessControlType.Deny);
                if (preset)
                {
                    directoryEntry.ObjectSecurity.AddAccessRule(rule);
                }
                else
                {
                    directoryEntry.ObjectSecurity.RemoveAccessRule(rule);
                }
                directoryEntry.CommitChanges();
            }
            catch (Exception ex)
            {
                // log the error
            }
            return result;
        }
        #endregion

        #region  设置用户是否启用
        public void SetEnableAccount(DirectoryEntry dey, bool enable)
        {
            if (enable == true)
            {
                //启用账户
                //UF_DONT_EXPIRE_PASSWD 0x10000
                int exp = (int)dey.Properties["userAccountControl"].Value;
                dey.Properties["userAccountControl"].Value = exp | 0x0001;
                dey.CommitChanges();
                //UF_ACCOUNTDISABLE 0x0002
                int val = (int)dey.Properties["userAccountControl"].Value;
                dey.Properties["userAccountControl"].Value = val & ~0x0002;
                dey.CommitChanges();
            }
            else
            {
                //禁用账户
                int val = (int)dey.Properties["userAccountControl"].Value;
                dey.Properties["userAccountControl"].Value = val | 0x0002;
                dey.Properties["msExchHideFromAddressLists"].Value = "TRUE";
                dey.CommitChanges();
                dey.Close();
            }

        }
        #endregion

        #region 随机生成字符串（数字和字母混和）
        public string GenerateCheckCode(int codeCount)
        {
            var j = 0;
            string str = string.Empty;
            long num2 = DateTime.Now.Ticks + j;
            j++;
            Random random = new Random(((int)(((ulong)num2) & 0xffffffffL)) | ((int)(num2 >> j)));
            for (int i = 0; i < codeCount; i++)
            {
                char ch;
                int num = random.Next();
                if ((num % 2) == 0)
                {
                    ch = (char)(0x30 + ((ushort)(num % 10)));
                }
                else
                {
                    ch = (char)(0x41 + ((ushort)(num % 0x1a)));
                }
                str = str + ch.ToString();
            }
            return str;
        }

        #endregion

        #region 给用户发送短信

        /// <summary>
        /// 2016年9月开始使用的短信服务商
        /// </summary>
        /// <param name="mobs"></param>
        /// <param name="randonPwd"></param>
        /// <returns></returns>
        public string SendMsg(string mobile, string randonPwd)
        {
            bool msgResult = false;
            string str = String.Empty;
            AdmgrModel admgrModel = new AdmgrModel();
            try
            {
                var systemInfo = admgrModel.AdmgrSystems.FirstOrDefault(a => a.AdmgrSystemId == 1);
                #region 自己的代码

                //string Msg_Url = systemInfo.MsgUrl;
                //string Msg_Account = systemInfo.MsgAccount; 
                //string Msg_Pwd = systemInfo.MsgPwd; 
                ////string name = "888";
                //string msg = "您于" + DateTime.Now.ToString("yyyy-MM-dd HH:mm") + "重置了密码，您的新密码是" + randonPwd;
                //StringBuilder arge = new StringBuilder();
                //arge.AppendFormat("account={0}", Msg_Account);
                //arge.AppendFormat("&pswd={0}", Msg_Pwd);
                //arge.AppendFormat("&mobile={0}", mobile);
                //arge.AppendFormat("&msg={0}", msg);
                //arge.AppendFormat("&needstatus=true");
                //arge.AppendFormat("&extno=");
                //#region 发送短信
                //byte[] byteArray = Encoding.UTF8.GetBytes(arge.ToString());
                //HttpWebRequest webRequest = (HttpWebRequest)WebRequest.Create(new Uri(Msg_Url));
                //webRequest.Method = "POST";
                //webRequest.ContentType = "application/x-www-form-urlencoded";
                //webRequest.ContentLength = byteArray.Length;
                //Stream newStream = webRequest.GetRequestStream();
                //newStream.Write(byteArray, 0, byteArray.Length);
                //newStream.Close();
                ////接收返回信息：
                //HttpWebResponse response = (HttpWebResponse)webRequest.GetResponse();
                //if (response.StatusCode == HttpStatusCode.OK)
                //{
                //    StreamReader reader = new StreamReader(response.GetResponseStream(), Encoding.UTF8);
                //    str = reader.ReadToEnd();
                //    string[] str2 = str.Split(',');
                //    if (str2[1] == "0")
                //    {
                //        //成功
                //        msgResult = true;
                //    }
                //}
                #endregion

                #region 官方的代码
                string account = systemInfo.MsgAccount;
                string password = systemInfo.MsgPwd;
                string content = "您于" + DateTime.Now.ToString("yyyy-MM-dd HH:mm") + "重置了密码，您的新密码是" + randonPwd;
                string postStrTpl = "account={0}&pswd={1}&mobile={2}&msg={3}&needstatus=true&extno=";

                UTF8Encoding encoding = new UTF8Encoding();
                byte[] postData = encoding.GetBytes(string.Format(postStrTpl, account, password, mobile, content));

                HttpWebRequest myRequest = (HttpWebRequest)WebRequest.Create(systemInfo.MsgUrl);
                myRequest.Method = "POST";
                myRequest.ContentType = "application/x-www-form-urlencoded";
                myRequest.ContentLength = postData.Length;

                Stream newStream = myRequest.GetRequestStream();
                // Send the data.
                newStream.Write(postData, 0, postData.Length);
                newStream.Flush();
                newStream.Close();

                HttpWebResponse myResponse = (HttpWebResponse)myRequest.GetResponse();
                if (myResponse.StatusCode == HttpStatusCode.OK)
                {
                    StreamReader reader = new StreamReader(myResponse.GetResponseStream(), Encoding.UTF8);
                    str = reader.ReadToEnd();
                    //反序列化upfileMmsMsg.Text
                    //实现自己的逻辑
                }
                else
                {
                    //访问失败
                    str = "访问失败";
                }

                #endregion
            }
            catch
            {
            }
            finally
            {
                admgrModel.Dispose();
            }
            return str;
            //提交失败，可能余额不足，或者敏感词汇等等
        }
        #endregion

        #region ADtree数据生成

        #region 查询树下的数据，迭代循环

        public static string TreeSearch(Int64 id, DirectoryEntry user, StringBuilder ss)
        {
            string st = @",{id:" + (id) + ",pId:" + id / 10000 + ",name:'" + user.Name.Substring(3) + "',icon:'../../Content/img/parent.png',open:true}";
            ss.Append(st);
            id = id * 10000;
            foreach (DirectoryEntry item in user.Children)
            {
                if (item.Properties["objectClass"][1].ToString() == "organizationalUnit")
                {
                    TreeSearch(++id, item, ss);
                }
                else if (item.Properties["objectClass"][1].ToString() == "user" || item.Properties["objectClass"][1].ToString() == "person")
                {
                    string str = ",{id:" + (++id) + ",pId:" + id / 10000 + ",name:'" + item.Name.Substring(3) + "',icon:'../../Content/img/leaf.png'}";
                    ss.Append(str);
                }
                //不需要group数据显示出来
                //else if(item.Properties["objectClass"][1].ToString() == "group")
                //{
                //    string str = ",{id:" + (++id) + ",pId:" + id / 10000 + ",name:'" + item.Name + "',icon:'../../Content/img/group.png'}";
                //    ss.Append(str);
                //}
            }
            return ss.ToString();
        }

        #endregion

        #region 查询树下的数据，迭代循环，且不显示CN，仅显示OU

        public static string OuTreeSearch(Int64 id, DirectoryEntry user, StringBuilder ss)
        {
            string st = @",{id:" + (id) + ",pId:" + id / 10000 + ",name:'" + user.Name.Substring(3) + "',icon:'../../Content/img/parent.png',open:true}";
            ss.Append(st);
            id = id * 10000;
            foreach (DirectoryEntry item in user.Children)
            {
                if (item.Properties["objectClass"][1].ToString() == "organizationalUnit")
                {
                    OuTreeSearch(++id, item, ss);
                }
                //else if (item.Properties["objectClass"][1].ToString() == "user" || item.Properties["objectClass"][1].ToString() == "person")
                //{
                //    string str = ",{id:" + (++id) + ",pId:" + id / 10000 + ",name:'" + item.Name + "',icon:'../../Content/img/leaf.png'}";
                //    ss.Append(str);
                //}
                //else if (item.Properties["objectClass"][1].ToString() == "group")
                //{
                //    string str = ",{id:" + (++id) + ",pId:" + id / 10000 + ",name:'" + item.Name + "',icon:'../../Content/img/group.png'}";
                //    ss.Append(str);
                //}
            }
            return ss.ToString();
        }

        #endregion

        #region 查询树下的数据，不迭代循环

        public static string RootTreeSearch(Int64 id, DirectoryEntry user, StringBuilder ss)
        {
            string st = @",{id:" + (id) + ",pId:" + id / 10000 + ",name:'" + user.Name.Substring(3) + "',open:true}";
            if (id < 100010001)
            {
                ss.Append(st);
            }
            id = id * 10000;
            foreach (DirectoryEntry item in user.Children)
            {
                if (item.Name.Contains("OU"))
                {
                    RootTreeSearch(++id, item, ss);
                }
            }
            return ss.ToString();
        }
        #endregion

        #endregion

        #region 根据id和name查询元素下的元素
        // 首次访问页面时调用此函数，返回一个选定组下的tree的json数据
        // 注意，这里的参数name的定义是写在js代码里面的，这个后面可能会优化修改
        public string SearchItemToJsonString(Int64 id, DirectoryEntry entry)
        {
            //搜索当前目录下的组织和用户
            StringBuilder str = new StringBuilder();
            string return_str = "[" + TreeSearch(id, entry, str).Substring(1) + "]";
            return return_str;
        }
        public string SearchRootToJsonString(Int64 id, DirectoryEntry entry)
        {
            //搜索当前目录下的组织和用户
            StringBuilder str = new StringBuilder();
            string return_str = "[" + RootTreeSearch(id, entry, str).Substring(1) + "]";
            return return_str;
        }

        public string SearchEntryToJsonString(Int64 id, DirectoryEntry entry)
        {
            //搜索当前目录下的组织和用户
            StringBuilder str = new StringBuilder();
            string return_str = TreeSearch(id, entry, str).Substring(1);
            return return_str;
        }
        public string SearchOuEntryToJsonString(Int64 id, DirectoryEntry entry)
        {
            //搜索当前目录下的组织和用户
            StringBuilder str = new StringBuilder();
            string return_str = OuTreeSearch(id, entry, str).Substring(1);
            return return_str;
        }
        public string FindDirectoryEntryByName(Int64 id, DirectoryEntry entry, string name)
        {
            name = name.Substring(name.IndexOf("=") + 1);
            DirectorySearcher search = new DirectorySearcher(entry);
            search.Filter = "(name= " + name + ")";
            SearchResult resutl = search.FindOne();
            DirectoryEntry new_entry = resutl.GetDirectoryEntry();
            string zNodes = "," + SearchEntryToJsonString(id, new_entry);
            return zNodes;
        }
        public string FindOuDirectoryEntryByName(Int64 id, DirectoryEntry entry, string name)
        {
            name = name.Substring(name.IndexOf("=") + 1);
            DirectorySearcher search = new DirectorySearcher(entry);
            search.Filter = "(name= " + name + ")";
            SearchResult resutl = search.FindOne();
            DirectoryEntry new_entry = resutl.GetDirectoryEntry();
            string zNodes = "," + SearchOuEntryToJsonString(id, new_entry);
            return zNodes;
        }
        public DirectoryEntry FindCnDirectoryEntryByName(string name)
        {
            name = name.Substring(name.IndexOf("=") + 1);
            DirectorySearcher search = new DirectorySearcher(_root);
            search.Filter = "(sAMAccountName= " + name + ")";
            DirectoryEntry subEntry;
            try
            {
                subEntry = search.FindOne().GetDirectoryEntry();
            }
            catch (Exception e)
            {
                subEntry = null;
            }

            return subEntry;
        }
        #endregion

        #endregion
    }
}