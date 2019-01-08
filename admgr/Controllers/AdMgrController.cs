using admgr.EfModel;
using admgr.Entity;
using admgr.Models;
using ADmgr.Helper;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.DirectoryServices;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Web;
using System.Web.Mvc;
using System.Web.Security;

namespace admgr.Controllers
{
    public class AdMgrController : Controller
    {
        readonly AdmgrAd _admgrAd;

        public AdMgrController()
        {
            _admgrAd = new AdmgrAd();
        }

        #region 登录页面
        public ActionResult Login()
        {
            return View();
        }

        [HttpPost]
        public ActionResult LoginEdit()
        {
            string number = Request["name"].ToString();
            string password = Request["password"].ToString();
            string accountType = Request["accountType"].ToString();
            bool remember = Request["remember"] == "true";
            string errorStr = "登录失败";
            var dataJson = string.Empty;
            //roleNum:1为普通域用户，2为ad公司管理员，3为Exhcange管理员，4为Lync管理员，5为admgr系统管理员,6为集团公司管理员
            int roleNum = 1;
            try
            {
                bool codeVerifyResult = Request["vcode"] == Session["vCode"].ToString();
                if (!codeVerifyResult)
                {
                    errorStr = "验证码错误";
                    throw new Exception("");
                }

                AdmgrModel admgrModel = new AdmgrModel();
                //如果是域用户
                if (accountType == "1" )
                {
                        DirectoryEntry userEntry = _admgrAd.CheckLogin(number, password);
                        if (userEntry == null) throw new Exception("用户不存在");
                        //将用户的DirectoryEntry存入缓存
                        Session["userEntry"] = userEntry;
                        //验证成功后进行权限的赋值
                        var user = admgrModel.AdUsers.SingleOrDefault(a => a.Number == number);
                        if (user != null)
                        {
                            roleNum = user.Role.RoleNum;
                        }
                }
                else
                {
                    //进行管理员的验证
                    var pwd = HashUtility.GetSHA256Hash(password);
                    var adAdmin = admgrModel.Users.SingleOrDefault(a => a.UserName == number && a.UserPassword == pwd);
                    //验证成功后进行权限的赋值
                    roleNum = adAdmin.Role.RoleNum;
                    //将 roleNum写入MVC的role判断
                    errorStr = "登录成功";
                }
                Session["userName"] = number;
                Session["roleNum"] = roleNum;
                #region 将相关信息写入cookie
                var cookieJson = JsonConvert.SerializeObject(new { name = number, roleNum = roleNum });
                HttpCookie cookie = new HttpCookie("user");
                cookie.Value = cookieJson;
                if (remember)
                {
                    cookie.Expires = DateTime.Now.AddDays(7);
                }
                else
                {
                    cookie.Expires = DateTime.Now.AddHours(1);
                }
                Response.AppendCookie(cookie);
                #endregion
                dataJson = JsonConvert.SerializeObject(new { result = true,error = errorStr });
                admgrModel.Dispose();
            }
            catch (Exception ex)
            {
                dataJson = JsonConvert.SerializeObject(new { result = false, error = errorStr });
            }
            return Json(dataJson, JsonRequestBehavior.AllowGet);
        }
        #endregion

        #region 网站首页
        public ActionResult Index()
        {
            return View();
        }
        #endregion

        #region 忘记密码页面
        public ActionResult ForgetPassword()
        {
            return View();
        }
        [HttpPost]
        public ActionResult ForgetPasswordEdit()
        {
            var dataJson = string.Empty;
            try
            {
                var name = Request["name"];
                var mobile = Request["mobile"];
                string resultTemp = _admgrAd.ResetPwdByNameAndMobile(name, mobile);
                if (resultTemp != "访问失败")
                {
                    resultTemp = "修改成功";
                    dataJson = JsonConvert.SerializeObject(new { result = true, info = resultTemp });
                }
                else
                {
                    dataJson = JsonConvert.SerializeObject(new { result = false, info = resultTemp });
                }
             
            }
            catch
            {

            }
         
            return Json(dataJson, JsonRequestBehavior.AllowGet);
        }
        #endregion

        #region 个人用户信息页
        public ActionResult UserIndex()
        {
            try
            {
                string userName = Session["userName"].ToString();
                DirectoryEntry userEntry = (DirectoryEntry)Session["userEntry"];
                //取得用户信息
                AdUser adUser = _admgrAd.GetAdUserByAdName(userName);
                #region 读取用户头像
                byte[] imgData = _admgrAd.GetAvatarPath(userEntry);
                var relatvePath = "/Content/Photo/Avatar/DefaultUserPhoto.jpg";
                if (imgData != null)
                {
                    System.IO.MemoryStream ms = new System.IO.MemoryStream(imgData);
                    System.Drawing.Image img = System.Drawing.Image.FromStream(ms);
                    img.Save(Urlconvertorlocal(relatvePath));
                    relatvePath = "/Content/photo/Avatar/" + DateTime.Now.ToString("ddHHmmssff") + ".jpg";
                }
                #endregion
                ViewData["userImage"] = relatvePath;
                ViewData["number"] = adUser.Number;
                ViewData["adName"] = adUser.AdName;
                ViewData["sn"] = adUser.Sn;
                ViewData["displayName"] = adUser.DisplayName;
                ViewData["title"] = adUser.Title;
                ViewData["company"] = adUser.Company;
                ViewData["dept"] = adUser.Dept;
                ViewData["office"] = adUser.Office;
                ViewData["officephone"] = adUser.Officephone;
                ViewData["ext"] = adUser.Ext;
                ViewData["fax"] = adUser.Fax;
                ViewData["email"] = adUser.Email;
                ViewData["mobile"] = adUser.Mobile;
                ViewData["homephone"] = adUser.Homephone;
                ViewData["country"] = adUser.Country;
                ViewData["province"] = adUser.Province;
                ViewData["city"] = adUser.City;
                ViewData["zipcode"] = adUser.Zipcode;
                ViewData["address"] = adUser.Address;
                ViewData["description"] = adUser.Description;
                return View();
            }
            catch
            {
                return RedirectToAction("Login");
            }
        }

        [HttpPost]
        //用户编辑个人信息提交
        public ActionResult UserIndexEdit()
        {
            DirectoryEntry userEntry = (DirectoryEntry)Session["userEntry"];
            AdUser adUser = new AdUser();
            adUser.Office = Request["office"];
            adUser.Officephone = Request["officephone"];
            adUser.Ext = Request["ext"];
            adUser.Fax = Request["fax"];
            adUser.Email = Request["email"];
            adUser.Mobile = Request["mobile"];
            adUser.Homephone = Request["homephone"];
            adUser.Country = Request["country"];
            adUser.Province = Request["province"];
            adUser.City = Request["city"];
            adUser.Zipcode = Request["zipcode"];
            adUser.Address = Request["address"];
            adUser.Description = Request["description"];
            bool resultTemp = _admgrAd.SetAdUserInfo(userEntry, adUser);
            var dataJson = JsonConvert.SerializeObject(new { result = resultTemp });
            return Json(dataJson, JsonRequestBehavior.AllowGet);
        }

        #endregion

        #region 域用户管理页

        #region 公司域用户管理
        public ActionResult ManageCompanyAdUser()
        {
            return View();
        }

        //查询所有用户
        public ActionResult CompanyAdUserList()
        {
            string dataJson = String.Empty;
            try
            {
                //根据用户所在的公司，返回公司的所有的用户
                //目前的ad的规则是，用户在公司-部门-用户中的第三级，所以用户的上两级是公司的层，取得公司层的所有用户
                DirectoryEntry userEntry = (DirectoryEntry)Session["userEntry"];
                List<AdUser> adUserList = _admgrAd.GetCompanyAdUsers(userEntry);
                List<AdUserPartViewModel> adUserPartViewModelList = new List<AdUserPartViewModel>();
                #region 取得缓存中的数据
                List<CompanyViewModel> companyViewModels = new List<CompanyViewModel>();
                List<DeptViewModel> deptViewModels = new List<DeptViewModel>();
                List<TitleViewModel> titleViewModels = new List<TitleViewModel>();
                if (Session["companyViewModels"] != null && Session["deptViewModels"] != null && Session["titleViewModels"] != null)
                {
                    companyViewModels = (List<CompanyViewModel>)Session["companyViewModels"];
                    deptViewModels = (List<DeptViewModel>)Session["deptViewModels"];
                    titleViewModels = (List<TitleViewModel>)Session["titleViewModels"];

                }
                #endregion
                int i = 1;
                foreach (var item in adUserList)
                {
                    DirectoryEntry itemEntry = _admgrAd.GetDirectoryEntryByAdName(item.AdName);
                    var companyName = itemEntry.Parent.Parent.Properties["name"][0].ToString();
                    var deptName = itemEntry.Parent.Properties["name"][0].ToString();
                    AdUserPartViewModel adUserPartViewModel = new AdUserPartViewModel();
                    adUserPartViewModel.AdUserInfoId = (i++).ToString();
                    adUserPartViewModel.Sn = item.Sn;
                    adUserPartViewModel.Number = item.Number;
                    adUserPartViewModel.AdName = item.AdName;
                    adUserPartViewModel.CompanyId = companyViewModels.FirstOrDefault(a => a.CompanyName == companyName).CompanyId ;
                    adUserPartViewModel.DeptId = deptViewModels.FirstOrDefault(a => a.DeptName == deptName).DeptId ;
                    adUserPartViewModel.TitleId = titleViewModels.FirstOrDefault(a => a.TitleName == item.Title).TitleId ;
                    //adUserPartViewModel.DisplayName = item.DisplayName;
                    adUserPartViewModel.Mobile = item.Mobile;
                    adUserPartViewModel.Email = item.Email;
                    adUserPartViewModel.DisbleId = item.Disble == false ? "1" : "2";
                    adUserPartViewModelList.Add(adUserPartViewModel);
                }
                Session["adUserPartViewModels"] = adUserPartViewModelList;
                if (adUserList != null)
                {
                    dataJson = JsonConvert.SerializeObject(adUserPartViewModelList);
                }
            }
            catch (Exception ex)
            {

            }
            return JavaScript(dataJson);
        }
        //创建多个用户
        public ActionResult CompanyAdUserAdd()
        {
            var dataJson = string.Empty;
            try
            {
                #region 取得缓存中的数据
                List<CompanyViewModel> companyViewModels = new List<CompanyViewModel>();
                List<DeptViewModel> deptViewModels = new List<DeptViewModel>();
                List<TitleViewModel> titleViewModels = new List<TitleViewModel>();
                if (Session["companyViewModels"] != null && Session["deptViewModels"] != null && Session["titleViewModels"] != null)
                {
                    companyViewModels = (List<CompanyViewModel>)Session["companyViewModels"];
                    deptViewModels = (List<DeptViewModel>)Session["deptViewModels"];
                    titleViewModels = (List<TitleViewModel>)Session["titleViewModels"];

                }
                #endregion
                string models = System.Web.HttpContext.Current.Request["models"];
                var modelsJsonArray = JsonConvert.DeserializeObject<JArray>(models);
                AdUserPart adUserPart = new AdUserPart();
                adUserPart.Sn = modelsJsonArray[0]["Sn"].ToString();
                adUserPart.Number = modelsJsonArray[0]["Number"].ToString();
                adUserPart.AdName = modelsJsonArray[0]["AdName"].ToString();
                adUserPart.Company = companyViewModels.FirstOrDefault(a => a.CompanyId == modelsJsonArray[0]["CompanyId"]["CompanyId"].ToString()).CompanyName;
                adUserPart.Dept = deptViewModels.FirstOrDefault(a => a.DeptId == modelsJsonArray[0]["DeptId"]["DeptId"].ToString()).DeptName;
                adUserPart.Title = titleViewModels.FirstOrDefault(a => a.TitleId == modelsJsonArray[0]["TitleId"]["TitleId"].ToString()).TitleName;
                //显示名自动生成
                adUserPart.DisplayName = adUserPart.Dept + adUserPart.Title.Substring(adUserPart.Title.Length - 1) + "-" + adUserPart.Sn;
                adUserPart.Mobile = modelsJsonArray[0]["Mobile"].ToString();
                adUserPart.Email = modelsJsonArray[0]["Email"].ToString();
                adUserPart.Disble = modelsJsonArray[0]["DisbleId"]["DisbleId"].ToString() == "2" ? false : true;
                DirectoryEntry userEntry = _admgrAd.CreateUser(adUserPart);
                AdUserPartViewModel adUserPartViewModel = new AdUserPartViewModel();
                adUserPartViewModel.Sn = adUserPart.Sn;
                adUserPartViewModel.Number = adUserPart.Number;
                adUserPartViewModel.AdName = adUserPart.AdName;
                adUserPartViewModel.CompanyId = modelsJsonArray[0]["CompanyId"]["CompanyId"].ToString();
                adUserPartViewModel.DeptId = modelsJsonArray[0]["DeptId"]["DeptId"].ToString();
                adUserPartViewModel.TitleId = modelsJsonArray[0]["TitleId"]["TitleId"].ToString();
                adUserPartViewModel.Mobile = adUserPart.Mobile;
                adUserPartViewModel.Email = adUserPart.Email;
                adUserPartViewModel.DisbleId = modelsJsonArray[0]["DisbleId"]["DisbleId"].ToString();
                dataJson = JsonConvert.SerializeObject(adUserPartViewModel);
            }
            catch (Exception ex)
            {
            }
            return JavaScript(dataJson);

        }
        //更新多个用户
        public ActionResult CompanyAdUserEdit()
        {
            var dataJson = string.Empty;
            try
            {
                #region 取得缓存中的数据
                List<CompanyViewModel> companyViewModels = new List<CompanyViewModel>();
                List<DeptViewModel> deptViewModels = new List<DeptViewModel>();
                List<TitleViewModel> titleViewModels = new List<TitleViewModel>();
                List<AdUserPartViewModel> adUserPartViewModels = new List<AdUserPartViewModel>();
                if (Session["companyViewModels"] != null && Session["deptViewModels"] != null && Session["titleViewModels"] != null && Session["adUserPartViewModels"] != null)
                {
                    companyViewModels = (List<CompanyViewModel>)Session["companyViewModels"];
                    deptViewModels = (List<DeptViewModel>)Session["deptViewModels"];
                    titleViewModels = (List<TitleViewModel>)Session["titleViewModels"];
                    adUserPartViewModels = (List<AdUserPartViewModel>)Session["adUserPartViewModels"];
                }
                #endregion
                string models = System.Web.HttpContext.Current.Request["models"];
                var modelsJsonArray = JsonConvert.DeserializeObject<JArray>(models);
                var oldAdUserInfoId = modelsJsonArray[0]["AdUserInfoId"].ToString();
                var oldAdName = adUserPartViewModels.FirstOrDefault(a => a.AdUserInfoId == oldAdUserInfoId).AdName;
                DirectoryEntry userEntry = _admgrAd.GetDirectoryEntryByAdName(oldAdName);
                AdUserPart adUserPart = new AdUserPart();
                adUserPart.Sn = modelsJsonArray[0]["Sn"].ToString();
                adUserPart.Number = modelsJsonArray[0]["Number"].ToString();
                adUserPart.AdName = modelsJsonArray[0]["AdName"].ToString();
                adUserPart.Company = companyViewModels.FirstOrDefault(a => a.CompanyId == modelsJsonArray[0]["CompanyId"].ToString()).CompanyName;
                adUserPart.Dept = deptViewModels.FirstOrDefault(a => a.DeptId == modelsJsonArray[0]["DeptId"].ToString() && a.CompanyId == modelsJsonArray[0]["CompanyId"].ToString()).DeptName;
                adUserPart.Title = titleViewModels.FirstOrDefault(a => a.TitleId == modelsJsonArray[0]["TitleId"].ToString()).TitleName;
                //显示名自动生成
                adUserPart.DisplayName = adUserPart.Dept + adUserPart.Title.Substring(adUserPart.Title.Length - 1) + "-" + adUserPart.Sn;
                adUserPart.Mobile = modelsJsonArray[0]["Mobile"].ToString();
                adUserPart.Email = modelsJsonArray[0]["Email"].ToString();
                adUserPart.Disble = modelsJsonArray[0]["DisbleId"].ToString() == "2" ? false : true;
                bool result = _admgrAd.SetAdUserInfoPart(userEntry, adUserPart);

                if (result)
                {
                    #region 如果公司或者部门发生了变化，则移动用户
                    var oldComapanyName = userEntry.Parent.Parent.Properties["name"][0].ToString();
                    var oldDeptName = userEntry.Parent.Properties["name"][0].ToString();
                    if ((oldComapanyName != adUserPart.Company) || (oldDeptName != adUserPart.Dept))
                    {
                        result = _admgrAd.MoveUserFromDeptToDept(userEntry, oldDeptName, adUserPart.Dept);
                    }
                    #endregion
                    AdUserPartViewModel adUserPartViewModel = new AdUserPartViewModel();
                    adUserPartViewModel.Sn = adUserPart.Sn;
                    adUserPartViewModel.Number = adUserPart.Number;
                    adUserPartViewModel.AdName = adUserPart.AdName;
                    adUserPartViewModel.CompanyId = modelsJsonArray[0]["CompanyId"].ToString();
                    adUserPartViewModel.DeptId = modelsJsonArray[0]["DeptId"].ToString();
                    adUserPartViewModel.TitleId = modelsJsonArray[0]["TitleId"].ToString();
                    adUserPartViewModel.Mobile = adUserPart.Mobile;
                    adUserPartViewModel.Email = adUserPart.Email;
                    adUserPartViewModel.DisbleId = modelsJsonArray[0]["DisbleId"].ToString();
                    dataJson = JsonConvert.SerializeObject(adUserPartViewModel);
                }
            }
            catch
            {

            }
            return JavaScript(dataJson);

        }
        //删除用户到回收站目录
        public ActionResult CompanyAdUserDelete()
        {
            var dataJson = string.Empty;
            try
            {
                string models = System.Web.HttpContext.Current.Request["models"];
                var modelsJsonArray = JsonConvert.DeserializeObject<JArray>(models);
                string adName = modelsJsonArray[0]["AdName"].ToString();
                DirectoryEntry userEntry = _admgrAd.GetDirectoryEntryByAdName(adName);
                bool result = _admgrAd.DeleteAdUser(userEntry);
                if (result)
                {
                    dataJson = models;
                }
            }
            catch
            {

            }
            return JavaScript(dataJson);
        }
        #endregion

        #region 集团域用户管理
        public ActionResult ManageAdUser()
        {
            return View();
        }

        //查询所有用户
        public ActionResult AdUserList()
        {
            string dataJson = String.Empty;
            try
            {
                //查询所有的用户
                List<AdUser> adUserList = _admgrAd.GetAllAdUsers();
                List<AdUserPartViewModel> adUserPartViewModelList = new List<AdUserPartViewModel>();
                #region 取得缓存中的数据
                List<CompanyViewModel> companyViewModels = new List<CompanyViewModel>();
                List<DeptViewModel> deptViewModels = new List<DeptViewModel>();
                List<TitleViewModel> titleViewModels = new List<TitleViewModel>();
                if (Session["companyViewModels"] != null && Session["deptViewModels"] != null && Session["titleViewModels"] != null)
                {
                    companyViewModels = (List<CompanyViewModel>)Session["companyViewModels"];
                    deptViewModels = (List<DeptViewModel>)Session["deptViewModels"];
                    titleViewModels = (List<TitleViewModel>)Session["titleViewModels"];
                }
                #endregion
                int i = 1;
                foreach (var item in adUserList)
                {
                    DirectoryEntry userEntry = _admgrAd.GetDirectoryEntryByAdName(item.AdName);
                    var companyName = userEntry.Parent.Parent.Properties["name"][0].ToString();
                    var deptName = userEntry.Parent.Properties["name"][0].ToString();
                    AdUserPartViewModel adUserPartViewModel = new AdUserPartViewModel();
                    adUserPartViewModel.AdUserInfoId = (i++).ToString();
                    adUserPartViewModel.Sn = item.Sn;
                    adUserPartViewModel.Number = item.Number;
                    adUserPartViewModel.AdName = item.AdName;
                    adUserPartViewModel.CompanyId = companyViewModels.FirstOrDefault(a => a.CompanyName == companyName) == null ? "1" : companyViewModels.FirstOrDefault(a => a.CompanyName == companyName).CompanyId;
                    adUserPartViewModel.DeptId = deptViewModels.FirstOrDefault(a => a.DeptName == deptName) == null ? "1" : deptViewModels.FirstOrDefault(a => a.DeptName == deptName).DeptId;
                    adUserPartViewModel.TitleId = titleViewModels.FirstOrDefault(a => a.TitleName == item.Title) == null ? "1" : titleViewModels.FirstOrDefault(a => a.TitleName == item.Title).TitleId;
                    //adUserPartViewModel.DisplayName = item.DisplayName;
                    adUserPartViewModel.Mobile = item.Mobile;
                    adUserPartViewModel.Email = item.Email;
                    adUserPartViewModel.DisbleId = item.Disble == false ? "1" : "2";
                    adUserPartViewModelList.Add(adUserPartViewModel);
                }
                Session["adUserPartViewModels"] = adUserPartViewModelList;
                if (adUserList != null)
                {
                    dataJson = JsonConvert.SerializeObject(adUserPartViewModelList);
                }
            }
            catch (Exception ex)
            {

            }
            return JavaScript(dataJson);
        }
        //创建多个用户
        public ActionResult AdUserAdd()
        {
            var dataJson = string.Empty;
            try
            {
                #region 取得缓存中的数据
                List<CompanyViewModel> companyViewModels = new List<CompanyViewModel>();
                List<DeptViewModel> deptViewModels = new List<DeptViewModel>();
                List<TitleViewModel> titleViewModels = new List<TitleViewModel>();
                if (Session["companyViewModels"] != null && Session["deptViewModels"] != null && Session["titleViewModels"] != null)
                {
                    companyViewModels = (List<CompanyViewModel>)Session["companyViewModels"];
                    deptViewModels = (List<DeptViewModel>)Session["deptViewModels"];
                    titleViewModels = (List<TitleViewModel>)Session["titleViewModels"];

                }
                #endregion
                string models = System.Web.HttpContext.Current.Request["models"];
                var modelsJsonArray = JsonConvert.DeserializeObject<JArray>(models);
                AdUserPart adUserPart = new AdUserPart();
                adUserPart.Sn = modelsJsonArray[0]["Sn"].ToString();
                adUserPart.Number = modelsJsonArray[0]["Number"].ToString();
                adUserPart.AdName = modelsJsonArray[0]["AdName"].ToString();
                adUserPart.Company = companyViewModels.FirstOrDefault(a => a.CompanyId == modelsJsonArray[0]["CompanyId"]["CompanyId"].ToString()).CompanyName;
                adUserPart.Dept = deptViewModels.FirstOrDefault(a => a.DeptId == modelsJsonArray[0]["DeptId"]["DeptId"].ToString()).DeptName;
                adUserPart.Title = titleViewModels.FirstOrDefault(a => a.TitleId == modelsJsonArray[0]["TitleId"]["TitleId"].ToString()).TitleName;
                //显示名自动生成
                adUserPart.DisplayName = adUserPart.Dept + adUserPart.Title.Substring(adUserPart.Title.Length - 1) + "-" + adUserPart.Sn;
                adUserPart.Mobile = modelsJsonArray[0]["Mobile"].ToString();
                adUserPart.Email = modelsJsonArray[0]["Email"].ToString();
                adUserPart.Disble = modelsJsonArray[0]["DisbleId"]["DisbleId"].ToString() == "2" ? false : true;
                DirectoryEntry userEntry = _admgrAd.CreateUser(adUserPart);
                AdUserPartViewModel adUserPartViewModel = new AdUserPartViewModel();
                adUserPartViewModel.Sn = adUserPart.Sn;
                adUserPartViewModel.Number = adUserPart.Number;
                adUserPartViewModel.AdName = adUserPart.AdName;
                adUserPartViewModel.CompanyId = modelsJsonArray[0]["CompanyId"]["CompanyId"].ToString();
                adUserPartViewModel.DeptId = modelsJsonArray[0]["DeptId"]["DeptId"].ToString();
                adUserPartViewModel.TitleId = modelsJsonArray[0]["TitleId"]["TitleId"].ToString();
                adUserPartViewModel.Mobile = adUserPart.Mobile;
                adUserPartViewModel.Email = adUserPart.Email;
                adUserPartViewModel.DisbleId = modelsJsonArray[0]["DisbleId"]["DisbleId"].ToString();
                dataJson = JsonConvert.SerializeObject(adUserPartViewModel);
            }
            catch (Exception ex)
            {
            }
            return JavaScript(dataJson);

        }
        //更新多个用户
        public ActionResult AdUserEdit()
        {
            var dataJson = string.Empty;
            try
            {
                #region 取得缓存中的数据
                List<CompanyViewModel> companyViewModels = new List<CompanyViewModel>();
                List<DeptViewModel> deptViewModels = new List<DeptViewModel>();
                List<TitleViewModel> titleViewModels = new List<TitleViewModel>();
                List<AdUserPartViewModel> adUserPartViewModels = new List<AdUserPartViewModel>();
                if (Session["companyViewModels"] != null && Session["deptViewModels"] != null && Session["titleViewModels"] != null && Session["adUserPartViewModels"] != null)
                {
                    companyViewModels = (List<CompanyViewModel>)Session["companyViewModels"];
                    deptViewModels = (List<DeptViewModel>)Session["deptViewModels"];
                    titleViewModels = (List<TitleViewModel>)Session["titleViewModels"];
                    adUserPartViewModels = (List<AdUserPartViewModel>)Session["adUserPartViewModels"];
                }
                #endregion
                string models = System.Web.HttpContext.Current.Request["models"];
                var modelsJsonArray = JsonConvert.DeserializeObject<JArray>(models);
                var oldAdUserInfoId = modelsJsonArray[0]["AdUserInfoId"].ToString();
                var oldAdName = adUserPartViewModels.FirstOrDefault(a => a.AdUserInfoId == oldAdUserInfoId).AdName;
                DirectoryEntry userEntry = _admgrAd.GetDirectoryEntryByAdName(oldAdName);
                AdUserPart adUserPart = new AdUserPart();
                adUserPart.Sn = modelsJsonArray[0]["Sn"].ToString();
                adUserPart.Number = modelsJsonArray[0]["Number"].ToString();
                adUserPart.AdName = modelsJsonArray[0]["AdName"].ToString();
                adUserPart.Company = companyViewModels.FirstOrDefault(a => a.CompanyId == modelsJsonArray[0]["CompanyId"].ToString()).CompanyName;
                adUserPart.Dept = deptViewModels.FirstOrDefault(a => a.DeptId == modelsJsonArray[0]["DeptId"].ToString() && a.CompanyId == modelsJsonArray[0]["CompanyId"].ToString()).DeptName;
                adUserPart.Title = titleViewModels.FirstOrDefault(a => a.TitleId == modelsJsonArray[0]["TitleId"].ToString()).TitleName;
                //显示名自动生成
                adUserPart.DisplayName = adUserPart.Dept + adUserPart.Title.Substring(adUserPart.Title.Length - 1) + "-" + adUserPart.Sn;
                adUserPart.Mobile = modelsJsonArray[0]["Mobile"].ToString();
                adUserPart.Email = modelsJsonArray[0]["Email"].ToString();
                adUserPart.Disble = modelsJsonArray[0]["DisbleId"].ToString() == "2" ? false : true;
                bool result = _admgrAd.SetAdUserInfoPart(userEntry, adUserPart);

                if (result)
                {
                    #region 如果公司或者部门发生了变化，则移动用户
                    var oldComapanyName = userEntry.Parent.Parent.Properties["name"][0].ToString();
                    var oldDeptName = userEntry.Parent.Properties["name"][0].ToString();
                    if ((oldComapanyName != adUserPart.Company) || (oldDeptName != adUserPart.Dept))
                    {
                        result = _admgrAd.MoveUserFromDeptToDept(userEntry, oldDeptName, adUserPart.Dept);
                    }
                    #endregion
                    AdUserPartViewModel adUserPartViewModel = new AdUserPartViewModel();
                    adUserPartViewModel.Sn = adUserPart.Sn;
                    adUserPartViewModel.Number = adUserPart.Number;
                    adUserPartViewModel.AdName = adUserPart.AdName;
                    adUserPartViewModel.CompanyId = modelsJsonArray[0]["CompanyId"].ToString();
                    adUserPartViewModel.DeptId = modelsJsonArray[0]["DeptId"].ToString();
                    adUserPartViewModel.TitleId = modelsJsonArray[0]["TitleId"].ToString();
                    adUserPartViewModel.Mobile = adUserPart.Mobile;
                    adUserPartViewModel.Email = adUserPart.Email;
                    adUserPartViewModel.DisbleId = modelsJsonArray[0]["DisbleId"].ToString();
                    dataJson = JsonConvert.SerializeObject(adUserPartViewModel);
                }
            }
            catch
            {

            }
            return JavaScript(dataJson);

        }
        //删除用户到回收站目录
        public ActionResult AdUserDelete()
        {
            var dataJson = string.Empty;
            try
            {
                string models = System.Web.HttpContext.Current.Request["models"];
                var modelsJsonArray = JsonConvert.DeserializeObject<JArray>(models);
                string adName = modelsJsonArray[0]["AdName"].ToString();
                DirectoryEntry userEntry = _admgrAd.GetDirectoryEntryByAdName(adName);
                bool result = _admgrAd.DeleteAdUser(userEntry);
                if (result)
                {
                    dataJson = models;
                }
            }
            catch
            {

            }
            return JavaScript(dataJson);
        }
        #endregion

        #endregion

        #region 域组织单位管理页
        public ActionResult ManageOrganizationUnit()
        {
            return View();
        }
        #region 公司的操作

        //查询ad下的所有公司
        public ActionResult GetCompanies()
        {
            string dataJson = String.Empty;
            try
            {
                List<CompanyViewModel> models = _admgrAd.GetCompanies();
                Session["onlyCompanyViewModels"] = models;
                dataJson = JsonConvert.SerializeObject(models);
            }
            catch
            {
                dataJson = JsonConvert.SerializeObject(new { error = "无法读取AD数据" });
            }
            return JavaScript(dataJson);
        }
        //查询ad下所有的公司和部门
        public ActionResult GetCompaniesAndDepts()
        {
            var dataJson = string.Empty;
            try
            {
                List<CompanyViewModel> companyViewModels = new List<CompanyViewModel>();
                List<DeptViewModel> deptViewModels = new List<DeptViewModel>();
                CompanyAndDeptViewModel companyAndDeptViewModel = _admgrAd.GetCompaniesAndDepts();
                companyViewModels = companyAndDeptViewModel.CompanyViewModelList;
                deptViewModels = companyAndDeptViewModel.DeptViewModelList;
                Session["companyViewModels"] = companyViewModels;
                Session["deptViewModels"] = deptViewModels;
                dataJson = JsonConvert.SerializeObject(companyAndDeptViewModel);
            }
            catch (Exception ex)
            {
                dataJson = JsonConvert.SerializeObject(new { error = "无法读取AD数据" });
            }
            return JavaScript(dataJson);


        }
        //查询ad下所有的公司和部门，只查询相关用户下的公司和部门
        public ActionResult GetCompaniesAndDeptsByUserEntry()
        {
            var dataJson = string.Empty;
            try
            {
                DirectoryEntry userEntry = (DirectoryEntry)Session["userEntry"];
                List<CompanyViewModel> companyViewModels = new List<CompanyViewModel>();
                List<DeptViewModel> deptViewModels = new List<DeptViewModel>();
                CompanyAndDeptViewModel companyAndDeptViewModel = _admgrAd.GetCompaniesAndDeptsByUserEntry(userEntry);
                companyViewModels = companyAndDeptViewModel.CompanyViewModelList;
                deptViewModels = companyAndDeptViewModel.DeptViewModelList;
                Session["companyViewModels"] = companyViewModels;
                Session["deptViewModels"] = deptViewModels;
                dataJson = JsonConvert.SerializeObject(companyAndDeptViewModel);
            }
            catch (Exception ex)
            {
                dataJson = JsonConvert.SerializeObject(new { error = "无法读取AD数据" });
            }
            return JavaScript(dataJson);


        }
        //增加公司
        public ActionResult AddCompany()
        {
            var dataJson = string.Empty;
            try
            {
                var companyId = Request.Form["CompanyId"].ToString();
                var companyNewName = Request.Form["CompanyName"].ToString();
                bool result = _admgrAd.AddCompany(companyNewName);
                if (result)
                {
                    #region 更新session
                    var companyViewModels = (List<CompanyViewModel>)Session["onlyCompanyViewModels"];
                    CompanyViewModel companyViewModel = new CompanyViewModel
                    {
                        CompanyId = companyId,
                        CompanyName = companyNewName
                    };
                    companyViewModels.Add(companyViewModel);
                    Session["onlyCompanyViewModels"] = companyViewModels;
                    #endregion

                    dataJson = JsonConvert.SerializeObject(new CompanyViewModel { CompanyId = companyId, CompanyName = companyNewName }); ;
                }
            }
            catch
            {

            }
            return JavaScript(dataJson);
        }
        //修改公司
        public ActionResult EditCompany()
        {
            string dataJson = String.Empty;
            try
            {
                var companyId = Request.Form["CompanyId"].ToString();
                var companyNewName = Request.Form["CompanyName"].ToString();
                List<CompanyViewModel> companyOnlyViewModels = (List<CompanyViewModel>)Session["onlyCompanyViewModels"];
                var companyOldName = companyOnlyViewModels.FirstOrDefault(a=>a.CompanyId == companyId).CompanyName;
                bool result = _admgrAd.EditCompany(companyOldName, companyNewName);
                if (result)
                {
                    #region 更新session
                    var companyViewModels = (List<CompanyViewModel>)Session["onlyCompanyViewModels"];
                    var model = companyViewModels.FirstOrDefault(a => a.CompanyId == companyId);
                    model.CompanyName = companyNewName;
                    Session["onlyCompanyViewModels"] = companyViewModels;
                    #endregion

                    dataJson = JsonConvert.SerializeObject(new CompanyViewModel { CompanyId = companyId, CompanyName = companyNewName });
                }
            }
            catch(Exception ex)
            {

            }
            finally
            {

            }
            return JavaScript(dataJson);
        }
        #endregion

        #region 部门的操作
        //根据公司id查询公司下的所有部门
        public ActionResult GetDepts()
        {
            string dataJson = String.Empty;
            try
            {
                var companyId = Request.Form[3].ToString();
                var companyViewModels = (List<CompanyViewModel>)Session["onlyCompanyViewModels"];
                var companyName = companyViewModels.FirstOrDefault(a => a.CompanyId == companyId).CompanyName;
                CompanyViewModel companyViewModel = new CompanyViewModel { CompanyId = companyId, CompanyName = companyName };
                List<DeptViewModel> deptViewModels = _admgrAd.GetDeptsByCompany(companyViewModel);
                List<CompanyContainDeptViewModel> companyContainDeptViewModelList = new List<Models.CompanyContainDeptViewModel>();
                CompanyContainDeptViewModel companyContainDeptViewModel = new CompanyContainDeptViewModel();
                companyContainDeptViewModel.CompanyId = companyId;
                companyContainDeptViewModel.DeptViewModelList = deptViewModels;
                //每次点击公司都把公司中的部门的数据重新刷新一遍
                if (Session["companyContainDeptViewModels"] != null)
                {
                    companyContainDeptViewModelList = (List<CompanyContainDeptViewModel>)Session["companyContainDeptViewModels"];
                    var model = companyContainDeptViewModelList.FirstOrDefault(a => a.CompanyId == companyId);
                    //如果原来存在数据
                    if (model != null)
                    {
                        
                        //删除掉原来的数据
                        companyContainDeptViewModelList.Remove(model);
                        //新增数据
                        companyContainDeptViewModelList.Add(companyContainDeptViewModel);
                    }
                    //如果原来不存在数据
                    else
                    {
                        companyContainDeptViewModelList.Add(companyContainDeptViewModel);
                    }
                }
                else
                {
                    companyContainDeptViewModelList.Add(companyContainDeptViewModel);
                }
                //然后再缓存到session中
                Session["companyContainDeptViewModels"] = companyContainDeptViewModelList;
                dataJson = JsonConvert.SerializeObject(deptViewModels);
            }
            catch
            {
                dataJson = JsonConvert.SerializeObject(new { error = "无法读取AD数据" });
            }
            return JavaScript(dataJson);
        }
        //增加部门
        public ActionResult AddDept()
        {
            string dataJson = String.Empty;
            try
            {
                var companyId = Request.Form["CompanyId"].ToString();
                var deptName = Request.Form["DeptName"].ToString();
                var companyViewModels = (List<CompanyViewModel>)Session["onlyCompanyViewModels"];
                var companyName = companyViewModels.FirstOrDefault(a => a.CompanyId == companyId).CompanyName;
                bool result = _admgrAd.AddDept(companyName, deptName);
                if (result)
                {
                    #region 更新session
                    List<CompanyContainDeptViewModel>  companyContainDeptViewModelList = (List<CompanyContainDeptViewModel>)Session["companyContainDeptViewModels"];
                    var model = companyContainDeptViewModelList.FirstOrDefault(a => a.CompanyId == companyId);
                    var deptNewId = model.DeptViewModelList.Count + 1;
                    DeptViewModel DeptViewModel = new DeptViewModel
                    {
                        CompanyId = companyId,
                        DeptId = deptNewId.ToString(),
                        DeptName = deptName
                    };
                    model.DeptViewModelList.Add(DeptViewModel);
                    Session["companyContainDeptViewModels"] = companyContainDeptViewModelList;
                    #endregion
                    dataJson = JsonConvert.SerializeObject(new DeptViewModel { CompanyId = companyId, DeptId = deptNewId.ToString(), DeptName = deptName });
                }
            }
            catch(Exception ex)
            {

            }
            return JavaScript(dataJson);
        }
        //修改部门
        public ActionResult EditDept()
        {
            string dataJson = String.Empty;
            try
            {
                var companyId = Request.Form["CompanyId"].ToString();
                var deptId = Request.Form["DeptId"].ToString();
                var deptNewName = Request.Form["DeptName"].ToString();
                var companyViewModels = (List<CompanyViewModel>)Session["onlyCompanyViewModels"];
                var companyName = companyViewModels.FirstOrDefault(a => a.CompanyId == companyId).CompanyName;
                List<CompanyContainDeptViewModel> companyContainDeptViewModel = (List<CompanyContainDeptViewModel>)Session["companyContainDeptViewModels"];
                var deptOldName = companyContainDeptViewModel.FirstOrDefault(a => a.CompanyId == companyId).DeptViewModelList.FirstOrDefault(a => a.DeptId == deptId).DeptName;
                bool result = _admgrAd.EditDept(companyName, deptOldName, deptNewName);
                if (result)
                {
                    #region 更新session
                    List<CompanyContainDeptViewModel> companyContainDeptViewModelList = (List<CompanyContainDeptViewModel>)Session["companyContainDeptViewModels"];
                    var model = companyContainDeptViewModelList.FirstOrDefault(a => a.CompanyId == companyId);
                    var deptModel = model.DeptViewModelList.FirstOrDefault(a => a.DeptId == deptId);
                    deptModel.DeptName = deptNewName;
                    Session["companyContainDeptViewModels"] = companyContainDeptViewModelList;
                    #endregion
                    dataJson = JsonConvert.SerializeObject(new DeptViewModel { CompanyId = companyId, DeptId = deptId, DeptName = deptNewName });
                }
            }
            catch
            {

            }
            finally
            {

            }
            return JavaScript(dataJson);
        }
        //删除部门
        #endregion

        #region 角色的操作
        public ActionResult GetRoles()
        {
            string dataJson = String.Empty;
            if (Session["roles"] == null)
            {
                try
                {
                    AdmgrModel admgrModel = new AdmgrModel();
                    var roles = admgrModel.Roles.ToList();
                    List<RoleViewModel> roleViewModels = new List<RoleViewModel>();
                    foreach (var item in roles)
                    {
                        RoleViewModel model = new RoleViewModel();
                        model.RoleId = item.RoleId.ToString();
                        model.RoleName = item.RoleName;
                        roleViewModels.Add(model);
                    }
                    dataJson = JsonConvert.SerializeObject(roleViewModels);
                    admgrModel.Dispose();
                    Session["roles"] = dataJson;
                }
                catch
                {

                    dataJson = JsonConvert.SerializeObject(new { error = "无法读取数据库数据" });
                }
            }
            else
            {
                dataJson = Session["roles"].ToString();
            }
            return JavaScript(dataJson);
        }

        #endregion

        #region 获取所有职务
        public ActionResult GetTitles()
        {
            string dataJson = String.Empty;
            try
            {
                List<TitleViewModel> models = new List<TitleViewModel>();
                if (Session["titleViewModels"] == null)
                {
                    try
                    {
                        models = _admgrAd.GetTitles();
                        Session["titleViewModels"] = models;
                    }
                    catch
                    {

                    }
                }
                else
                {
                    models = (List<TitleViewModel>)Session["titleViewModels"];
                }
                dataJson = JsonConvert.SerializeObject(models);
            }
            catch
            {
                dataJson = JsonConvert.SerializeObject(new { error = "无法读取数据库数据" });
            }

            return JavaScript(dataJson);
            
        }
        #endregion

        #endregion

        #region 批量导入页
        public ActionResult ImportUser()
        {
            return View();
        }

        [HttpPost]
        public ActionResult ImportUserByFile()
        {
            return View();
        }
        #endregion

        #region AD域控的配置页
        public ActionResult AdSystemIndex()
        {
            AdmgrModel admgrModel = new AdmgrModel();
            var admgrSystem = admgrModel.AdmgrSystems.FirstOrDefault(a => a.AdmgrSystemId == 1);
            ViewData["serverIp"] = admgrSystem.AdServerIp;
            ViewData["domain"] = admgrSystem.AdDomain;
            ViewData["account"] = admgrSystem.AdManagerUserName;
            ViewData["password"] = admgrSystem.AdManagerPwd;
            admgrModel.Dispose();
            return View();
        }
        [HttpPost]
        public ActionResult AdSystemIndexEdit()
        {
            var data = new { result = false };
            try
            {
                string serverIp = Request["serverIp"].ToString();
                string domain = Request["domain"].ToString();
                string account = Request["account"].ToString();
                string password = Request["password"].ToString();
                AdmgrModel admgrModel = new AdmgrModel();
                var admgrSystem = admgrModel.AdmgrSystems.FirstOrDefault(a => a.AdmgrSystemId == 1);
                admgrSystem.AdServerIp = serverIp;
                admgrSystem.AdDomain = domain;
                admgrSystem.AdManagerUserName = account;
                admgrSystem.AdManagerPwd = password;
                var i = admgrModel.SaveChanges();
                admgrModel.Dispose();
                if (i > 0)
                {
                    data = new { result = true };
                }
            }
            catch
            {
            }
            var dataJson = JsonConvert.SerializeObject(data);
            return Json(dataJson, JsonRequestBehavior.AllowGet);
        }
        #endregion

        #region AD域控权限配置页
        public ActionResult AdUserAuthority()
        {


            return View();
        }
        //获取所有公司部门的用户
        public ActionResult GetAllAdUsers()
        {
            string dataJson = String.Empty;
            try
            {

               
                List<AdUserViewModel> adUserViewModels = new List<AdUserViewModel>();
                if (Session["adUserViewModels"] == null)
                {
                    var adUsers = _admgrAd.GetAllAdUsers();
                    int i = 1;
                    foreach (var item in adUsers)
                    {
                        AdUserViewModel model = new AdUserViewModel();
                        model.AdUserId = i.ToString();
                        model.AdUserName = item.DisplayName;
                        adUserViewModels.Add(model);
                        ++i;
                    }
                    Session["adUserViewModels"] = adUserViewModels;
                    dataJson = JsonConvert.SerializeObject(adUserViewModels);

                }
                else
                { 
                    adUserViewModels = (List<AdUserViewModel>)Session["adUserViewModels"];
                    dataJson = JsonConvert.SerializeObject(adUserViewModels);
                }
            }
            catch
            {
                dataJson = JsonConvert.SerializeObject(new { error = "无法读AD数据" });
            }
            return JavaScript(dataJson);
        }

        //public ActionResult AdUserAuthorityList()
        //{
        //    List<UserRoleViewModel> userRoleViewModels = new List<UserRoleViewModel>()
        //    {
        //        new UserRoleViewModel
        //        {
        //            Id = "1",
        //            AdUserId = "1",
        //            RoleId = "2"
        //        }
        //    };
        //    var dataJson = JsonConvert.SerializeObject(userRoleViewModels);
        //    return JavaScript(dataJson);
        //}

        public ActionResult AdUserAuthorityList()
        {
            string dataJson = String.Empty;
            try
            {
                AdmgrModel admgrModel = new AdmgrModel();
                var adUsers = admgrModel.AdUsers.ToList();
                List<UserRoleViewModel> userRoleViewModels = new List<UserRoleViewModel>();
                #region 获取所有用户的list列表，便于查出用户的id
                List<AdUserViewModel> adUserViewModels = new List<AdUserViewModel>();
                if (Session["adUserViewModels"] != null)
                {
                    adUserViewModels = (List<AdUserViewModel>)Session["adUserViewModels"];
                }
                else
                {
                    var adUsersTemp = _admgrAd.GetAllAdUsers();
                    int j = 1;
                    foreach (var item in adUsers)
                    {
                        AdUserViewModel model = new AdUserViewModel();
                        model.AdUserId = j.ToString();
                        model.AdUserName = item.DisplayName;
                        adUserViewModels.Add(model);
                        ++j;
                    }
                }
                #endregion

                int i = 1;
                foreach (var item in adUsers)
                {
                    UserRoleViewModel model = new UserRoleViewModel();
                    model.Id = i.ToString();
                    //如果用户在AD中被删除了，就直接在数据库中将用户删除
                    var adUserModel = adUserViewModels.FirstOrDefault(a => a.AdUserName == item.AdName);
                    if (adUserModel != null)
                    {
                        model.AdUserId = adUserViewModels.FirstOrDefault(a => a.AdUserName == item.AdName).AdUserId;
                        model.RoleId = item.Role.RoleId.ToString();
                        userRoleViewModels.Add(model);
                        ++i;
                    }
                    else
                    {
                        admgrModel.AdUsers.Remove(item);
                    }
                   
                   

                }
                admgrModel.SaveChanges();
                dataJson = JsonConvert.SerializeObject(userRoleViewModels);
                admgrModel.Dispose();
            }
            catch
            {
                dataJson = JsonConvert.SerializeObject(new { error = "无法读AD数据" });
            }
            return JavaScript(dataJson);
        }
        public ActionResult AdUserAuthorityAdd()
        {
            var dataJson = string.Empty;
            try
            {
                AdmgrModel admgrModel = new AdmgrModel();
                string models = System.Web.HttpContext.Current.Request["models"];
                List<UserRoleViewModel> userRoleViewModels = JsonConvert.DeserializeObject<List<UserRoleViewModel>>(models);
                List<AdUser> adUserList = new List<AdUser>();
                List<AdUserViewModel> adUserViewModels = new List<AdUserViewModel>();
                foreach (var item in userRoleViewModels)
                {
                    string adUserId = item.AdUserId;
                    int roleId = Convert.ToInt32(item.RoleId);

                    if (Session["adUserViewModels"] != null)
                    {
                        adUserViewModels = (List<AdUserViewModel>)Session["adUserViewModels"];
                    }
                    var adUserName = adUserViewModels.FirstOrDefault(a => a.AdUserId == adUserId).AdUserName;
                    //AdUser存入数据库中的字段不需要那么多的字段，仅需要名字就够了
                    var role = admgrModel.Roles.FirstOrDefault(a => a.RoleId == roleId);

                    AdUser adUser = _admgrAd.GetAdUserByDisplayName(adUserName);
                    adUser.Role = role;
                    adUserList.Add(adUser);
                    //增加缓存中的字段
                    AdUserViewModel adUserViewModel = new AdUserViewModel();
                    adUserViewModel.AdUserId = (adUserViewModels.Count + 1).ToString();
                    adUserViewModel.AdUserName = adUserViewModel.AdUserName;
                    adUserViewModels.Add(adUserViewModel);
                }
                admgrModel.AdUsers.AddRange(adUserList);
                admgrModel.SaveChanges();
                admgrModel.Dispose();
                Session["adUserViewModels"] = adUserViewModels;
                dataJson = JsonConvert.SerializeObject(userRoleViewModels);
            }
            catch (Exception ex)
            {

            }
            return JavaScript(dataJson);
        }
        public ActionResult AdUserAuthorityEdit()
        {
            var dataJson = string.Empty;
            try
            {
                AdmgrModel admgrModel = new AdmgrModel();
                string models = System.Web.HttpContext.Current.Request["models"];
                List<UserRoleViewModel> userRoleViewModels = JsonConvert.DeserializeObject<List<UserRoleViewModel>>(models);
                List<AdUser> adUserList = new List<AdUser>();
                List<AdUserViewModel> adUserViewModels = new List<AdUserViewModel>();
                foreach (var item in userRoleViewModels)
                {
                    string adUserId = item.AdUserId;
                    int roleId = Convert.ToInt32(item.RoleId);

                    if (Session["adUserViewModels"] != null)
                    {
                        adUserViewModels = (List<AdUserViewModel>)Session["adUserViewModels"];
                    }
                    var adUserName = adUserViewModels.FirstOrDefault(a => a.AdUserId == adUserId).AdUserName;
                    //AdUser存入数据库中的字段不需要那么多的字段，仅需要名字就够了
                    var role = admgrModel.Roles.FirstOrDefault(a => a.RoleId == roleId);
                    var adUser = admgrModel.AdUsers.FirstOrDefault(a => a.AdName == adUserName);
                    adUser.Role = role;
                }
                admgrModel.SaveChanges();
                admgrModel.Dispose();
                dataJson = JsonConvert.SerializeObject(userRoleViewModels);
            }
            catch
            {

            }
            return JavaScript(dataJson);
        }
        public ActionResult AdUserAuthorityDelete()
        {
            var dataJson = string.Empty;
            try
            {
                AdmgrModel admgrModel = new AdmgrModel();
                string models = System.Web.HttpContext.Current.Request["models"];
                List<UserRoleViewModel> userRoleViewModels = JsonConvert.DeserializeObject<List<UserRoleViewModel>>(models);
                List<AdUser> adUserList = new List<AdUser>();
                List<AdUserViewModel> adUserViewModels = new List<AdUserViewModel>();
                foreach (var item in userRoleViewModels)
                {
                    string adUserId = item.AdUserId;
                    int roleId = Convert.ToInt32(item.RoleId);

                    if (Session["adUserViewModels"] != null)
                    {
                        adUserViewModels = (List<AdUserViewModel>)Session["adUserViewModels"];
                    }
                    var adUserName = adUserViewModels.FirstOrDefault(a => a.AdUserId == adUserId).AdUserName;
                    //AdUser存入数据库中的字段不需要那么多的字段，仅需要名字就够了
                    var role = admgrModel.Roles.FirstOrDefault(a => a.RoleId == roleId);
                    var adUser = admgrModel.AdUsers.FirstOrDefault(a => a.AdName == adUserName);
                    //将缓存中的数据也同步删除，不考虑数据链接异常的错误
                    adUserViewModels.Remove(adUserViewModels.FirstOrDefault(a => a.AdUserId == adUserId));
                    adUserList.Add(adUser);
                }
                admgrModel.AdUsers.RemoveRange(adUserList);
                var i = admgrModel.SaveChanges();
                admgrModel.Dispose();
                Session["adUserViewModels"] = adUserViewModels;
                dataJson = JsonConvert.SerializeObject(userRoleViewModels);
            }
            catch
            {

            }
            return JavaScript(dataJson);
        }

        #endregion

        #region admgr系统配置页
        public ActionResult AdmgrSystemIndex()
        {
            AdmgrModel admgrModel = new AdmgrModel();
            var admgrSystem = admgrModel.AdmgrSystems.FirstOrDefault(a => a.AdmgrSystemId == 1);
            ViewData["account"] = admgrSystem.SystemAccount;
            ViewData["password"] = admgrSystem.SystemPwd;
            admgrModel.Dispose();
            return View();
        }
        [HttpPost]
        public ActionResult AdmgrSystemIndexEdit()
        {
            var data = new { result = false };
            try
            {
                string account = Request["account"].ToString();
                string password = Request["password"].ToString();
                AdmgrModel admgrModel = new AdmgrModel();
                var admgrSystem = admgrModel.AdmgrSystems.FirstOrDefault(a => a.AdmgrSystemId == 1);
                admgrSystem.SystemAccount = account;
                admgrSystem.SystemPwd = password;
                var i = admgrModel.SaveChanges();
                admgrModel.Dispose();
                if (i > 0)
                {
                    data = new { result = true };
                }
            }
            catch
            {
            }
            var dataJson = JsonConvert.SerializeObject(data);
            return Json(dataJson, JsonRequestBehavior.AllowGet);
        }
        #endregion

        #region Exchange系统配置页
        public ActionResult ExchangeSystemIndex()
        {
            AdmgrModel admgrModel = new AdmgrModel();
            var admgrSystem = admgrModel.AdmgrSystems.FirstOrDefault(a => a.AdmgrSystemId == 1);
            ViewData["ExchangeServerIp"] = admgrSystem.SystemAccount;
            ViewData["ExchangeAccount"] = admgrSystem.SystemAccount;
            ViewData["ExchangePassword"] = admgrSystem.SystemAccount;
            admgrModel.Dispose();
            return View();
        }
        [HttpPost]
        public ActionResult ExchangeSystemIndexEdit()
        {
            var data = new { result = false };
            try
            {
                string exchangeServerIp = Request["exchangeServerIp"].ToString();
                string exchangeAccount = Request["exchangeAccount"].ToString();
                string exchangePassword = Request["exchangePassword"].ToString();
                AdmgrModel admgrModel = new AdmgrModel();
                var admgrSystem = admgrModel.AdmgrSystems.FirstOrDefault(a => a.AdmgrSystemId == 1);
                admgrSystem.ExchangIp = exchangeServerIp;
                admgrSystem.ExchangeUserName = exchangeAccount;
                admgrSystem.ExchangePwd = exchangePassword;
                var i = admgrModel.SaveChanges();
                admgrModel.Dispose();
                if (i > 0)
                {
                    data = new { result = true };
                }
            }
            catch
            {
            }
            var dataJson = JsonConvert.SerializeObject(data);
            return Json(dataJson, JsonRequestBehavior.AllowGet);
        }
        #endregion

        #region Exchange分子公司邮箱配置页
        public ActionResult ExchangeCompanyMail()
        {
            return View();
        }
        //展示数据
        public ActionResult ExchangeCompanyMailList()
        {
            var dataJson = string.Empty;
            try
            {
                AdmgrModel admgrModel = new AdmgrModel();
                var models = admgrModel.EmailDatas.ToList();
                List<CompanyViewModel> companyViewModels = (List<CompanyViewModel>)Session["onlyCompanyViewModels"];
                List<EmailDataViewModel> emailDataViewModels = new List<EmailDataViewModel>();
                foreach (var item in models)
                {
                    EmailDataViewModel model = new EmailDataViewModel();
                    model.EmailDataId = item.EmailDataId;
                    var companyTemp = companyViewModels.FirstOrDefault(a => a.CompanyName == item.CompanyName);
                    model.CompanyId = companyTemp == null ? "1" : companyTemp.CompanyId;
                    model.MainMailAddress = item.MainMailAddress;
                    model.SecondMailAddress = item.SecondMailAddress;
                    model.EmailStoreDb = item.EmailStoreDb;
                    emailDataViewModels.Add(model);
                }
                if (models.Count>0) dataJson = JsonConvert.SerializeObject(emailDataViewModels);
                admgrModel.Dispose();
            }
            catch
            {

            }
            return JavaScript(dataJson);
        }

        //新增数据
        public ActionResult ExchangeCompanyMailAdd()
        {
            var dataJson = string.Empty;
            try
            {
                AdmgrModel admgrModel = new AdmgrModel();
                var companyId = Request.Form["CompanyId"].ToString();
                var mainMailAddress = Request.Form["MainMailAddress"].ToString();
                var secondMailAddress = Request.Form["SecondMailAddress"].ToString();
                var emailStore = Request.Form["EmailStoreDb"].ToString();
                List<CompanyViewModel> companyViewModels = (List<CompanyViewModel>)Session["onlyCompanyViewModels"];
                var companyName = companyViewModels.FirstOrDefault(a => a.CompanyId == companyId).CompanyName;
                EmailData emailData = new EmailData
                {
                    CompanyName = companyName,
                    MainMailAddress = mainMailAddress,
                    SecondMailAddress = secondMailAddress,
                    EmailStoreDb = emailStore
                };
                admgrModel.EmailDatas.Add(emailData);
                var i = admgrModel.SaveChanges();
                if (i > 0)
                {
                    EmailDataViewModel model = new EmailDataViewModel();
                    model.CompanyId = companyId;
                    model.MainMailAddress = mainMailAddress;
                    model.SecondMailAddress = secondMailAddress;
                    model.EmailStoreDb = emailStore;
                    dataJson = JsonConvert.SerializeObject(model);
                }
            }
            catch
            {
             
            }
            return JavaScript(dataJson);

        }
        //更新数据
        public ActionResult ExchangeCompanyMailUpdate()
        {
            var dataJson = string.Empty;
            try
            {
                AdmgrModel admgrModel = new AdmgrModel();
                var emailDataId =Convert.ToInt32( Request.Form["EmailDataId"]);
                var companyId = Request.Form["CompanyId"].ToString();
                var mainMailAddress = Request.Form["MainMailAddress"].ToString();
                var secondMailAddress = Request.Form["SecondMailAddress"].ToString();
                var emailStore = Request.Form["EmailStoreDb"].ToString();
                List<CompanyViewModel> companyViewModels = (List<CompanyViewModel>)Session["onlyCompanyViewModels"];
                var companyName = companyViewModels.FirstOrDefault(a => a.CompanyId == companyId).CompanyName;
                var emailDataModel = admgrModel.EmailDatas.FirstOrDefault(a => a.EmailDataId == emailDataId);
                emailDataModel.CompanyName = companyName;
                emailDataModel.MainMailAddress = mainMailAddress;
                emailDataModel.SecondMailAddress = secondMailAddress;
                emailDataModel.EmailStoreDb = emailStore;
                var i = admgrModel.SaveChanges();
                if (i > 0)
                {
                    EmailDataViewModel model = new EmailDataViewModel();
                    model.EmailDataId = emailDataId;
                    model.CompanyId = companyId;
                    model.MainMailAddress = mainMailAddress;
                    model.SecondMailAddress = secondMailAddress;
                    model.EmailStoreDb = emailStore;
                    dataJson = JsonConvert.SerializeObject(model);
                }
            }
            catch
            {

            }
            return JavaScript(dataJson);
        }
        //删除数据
        public ActionResult ExchangeCompanyMailDelete()
        {
            var dataJson = string.Empty;
            try
            {
                AdmgrModel admgrModel = new AdmgrModel();
                DirectoryEntry userEntry = (DirectoryEntry)Session["userEntry"];
                string models = System.Web.HttpContext.Current.Request["models"];
                List<EmailData> emailDataList = JsonConvert.DeserializeObject<List<EmailData>>(models);
                var emailDataTemp = admgrModel.EmailDatas.FirstOrDefault(a => a.EmailDataId == emailDataList[0].EmailDataId);
                EmailData emailData = admgrModel.EmailDatas.Remove(emailDataTemp);
                var i = admgrModel.SaveChanges();
                admgrModel.Dispose();
                if (i > 0)
                {
                    dataJson = JsonConvert.SerializeObject(emailData);
                }
            }
            catch
            {

            }

            return JavaScript(dataJson);
        }
        #endregion

        #region Lync系统配置页
        public ActionResult LyncSystemConfig()
        {
            AdmgrModel admgrModel = new AdmgrModel();
            var admgrSystem = admgrModel.AdmgrSystems.FirstOrDefault(a => a.AdmgrSystemId == 1);
            ViewData["LyncServerIp"] = admgrSystem.SystemAccount;
            ViewData["LyncAccount"] = admgrSystem.SystemAccount;
            ViewData["LyncPassword"] = admgrSystem.SystemAccount;
            admgrModel.Dispose();
            return View();
        }

        public ActionResult LyncSystemConfigEdit()
        {
            {
                var data = new { result = false };
                try
                {
                    string lyncServerIp = Request["LyncServerIp"].ToString();
                    string lyncAccount = Request["LyncAccount"].ToString();
                    string lyncPassword = Request["LyncPassword"].ToString();
                    AdmgrModel admgrModel = new AdmgrModel();
                    var admgrSystem = admgrModel.AdmgrSystems.FirstOrDefault(a => a.AdmgrSystemId == 1);
                    admgrSystem.LyncIp = lyncServerIp;
                    admgrSystem.LyncUserName = lyncAccount;
                    admgrSystem.LyncPwd = lyncPassword;
                    var i = admgrModel.SaveChanges();
                    admgrModel.Dispose();
                    if (i > 0)
                    {
                        data = new { result = true };
                    }
                }
                catch
                {
                }
                var dataJson = JsonConvert.SerializeObject(data);
                return Json(dataJson, JsonRequestBehavior.AllowGet);
            }
        }
        #endregion

        #region SMS系统配置页
        public ActionResult SmsSystemIndex()
        {
            AdmgrModel admgrModel = new AdmgrModel();
            var admgrSystem = admgrModel.AdmgrSystems.FirstOrDefault(a => a.AdmgrSystemId == 1);
            ViewData["MsgUrl"] = admgrSystem.MsgUrl;
            ViewData["MsgAccount"] = admgrSystem.MsgAccount;
            ViewData["MsgPwd"] = admgrSystem.MsgPwd;
            admgrModel.Dispose();
            return View();
        }
        [HttpPost]
        public ActionResult SmsSystemIndexEdit()
        {
            var data = new { result = false };
            try
            {
                string msgUrl = Request["MsgUrl"].ToString();
                string msgAccount = Request["MsgAccount"].ToString();
                string msgPwd = Request["MsgPwd"].ToString();
                AdmgrModel admgrModel = new AdmgrModel();
                var admgrSystem = admgrModel.AdmgrSystems.FirstOrDefault(a => a.AdmgrSystemId == 1);
                admgrSystem.MsgUrl = msgUrl;
                admgrSystem.MsgAccount = msgAccount;
                admgrSystem.MsgPwd = msgPwd;
                var i = admgrModel.SaveChanges();
                admgrModel.Dispose();
                if (i > 0)
                {
                    data = new { result = true };
                }
            }
            catch
            {
            }
            var dataJson = JsonConvert.SerializeObject(data);
            return Json(dataJson, JsonRequestBehavior.AllowGet);
        }
        #endregion

        #region 头像操作

        #region 头像原图上传，返回图像相对路径

        /// <summary>
        /// 上传头像
        /// </summary>
        /// <param name="qqfile"></param>
        /// <returns></returns>
        [HttpPost]
        [AllowAnonymous]
        public ActionResult ProcessUpload(string qqfile)
        {
            try
            {
                //讲图片上传到服务器上的Content\img下的自定义目录中，以月份为单位保存图片
                string uploadPath = Server.MapPath("~") + @"Content\img\" + DateTime.Now.ToString("yyyyMM") + "\\";
                string imgName = DateTime.Now.ToString("ddHHmmssff");
                string imgType = qqfile.Substring(qqfile.LastIndexOf(".", StringComparison.Ordinal));

                #region 将图片保存在定义的目录中
                if (!Directory.Exists(uploadPath))
                {
                    Directory.CreateDirectory(uploadPath);
                }
                System.Web.HttpBrowserCapabilitiesBase browser = Request.Browser;
                string browType = browser.Browser.ToLower();
                if (browType == "ie") //IE特殊处理
                {
                    var httpPostedFileBase = Request.Files["qqfile"];
                    if (httpPostedFileBase != null)
                    {
                        string fileName = httpPostedFileBase.FileName;
                        imgType = fileName.Substring(fileName.LastIndexOf(".", StringComparison.Ordinal));
                    }
                    var postedFileBase = Request.Files["qqfile"];
                    if (postedFileBase != null)
                        using (var inputStream = postedFileBase.InputStream)
                        {
                            using (var flieStream = new FileStream(uploadPath + imgName + imgType, FileMode.Create))
                            {
                                // 保存上传图片
                                inputStream.CopyTo(flieStream);
                            }
                        }
                }
                else
                {
                    using (var inputStream = Request.InputStream)
                    {
                        using (var flieStream = new FileStream(uploadPath + imgName + imgType, FileMode.Create))
                        {
                            // 保存上传图片
                            inputStream.CopyTo(flieStream);
                        }
                    }
                }
                #endregion

                // 将大图上传到Image Server（直接设置一个临时文件夹保存就行了）, 得到大图 UploadedImgUrl 
                // 这步可根据实际情况修改
                string largeImgFullPath = uploadPath + imgName + imgType;
                //这里对上传的原图像进行了比例缩放，这里会对图像进行裁剪操作，具体的比例就是你自定义的比例
                Image newImg = ImgHandler.ZoomPicture(StreamHelper.ImagePath2Img(largeImgFullPath), 260, 378);
                newImg.Save(largeImgFullPath);
                // 讲绝对路径存入session留着给缩放的action使用
                Session["UploadedImgUrl"] = largeImgFullPath;
                // 将相对路径传到前台的cron.js动作里面
                string relUrl = "../" + Urlconvertor(largeImgFullPath);

                //Url.Content(rel_url);
                // 重新打开largeImgFullPath这个文件，并使用缩略后的文件将原文件覆盖，因为原文件本来也没什么用了
                return Json(new { success = true, message = relUrl }, "text/plain", Encoding.UTF8, JsonRequestBehavior.AllowGet);
            }
            catch (Exception ex)
            {
                return Json(new { fail = true, message = ex.Message }, "text/plain", Encoding.UTF8, JsonRequestBehavior.AllowGet);
            }
        }
        #endregion

        #region 头像裁剪并保存
        /// <summary>
        /// 保存缩略图
        /// </summary>
        /// <param name="form"></param>
        /// <returns></returns>
        [HttpPost]
        [AllowAnonymous]
        public ActionResult AvatarAction(FormCollection form)
        {
            try
            {
                int x = (int)Convert.ToDouble(form["x"]);
                int y = (int)Convert.ToDouble(form["y"]);
                int w = (int)Convert.ToDouble(form["w"]);
                int h = (int)Convert.ToDouble(form["h"]);

                // 先取得图片的绝对路径
                string absolutePath = Session["UploadedImgUrl"].ToString();
                //获取等比例缩放后的图片
                try
                {
                    // 根据坐标、宽高 裁剪头像,获得小图的路径，这个cutImagePath和absolutePatp其实是一样的，是绝对值
                    string cutImagePath = ImgHandler.CutAvatar(absolutePath, x, y, w, h);

                    //string path = string.Empty;

                    if (!string.IsNullOrEmpty(cutImagePath))
                    {
                        // 获取裁剪后图像的Url
                        //byte[] bytes = StreamHelper.Image2ByteWithPath(cutImagePath);
                        //string finalSmallImgPath = "";//DoUploadImageWS(bytes);

                        //保存Path
                        //path = ImageUrl + finalSmallImgPath;
                        //path = finalSmallImgPath;
                        //ViewBag.Path = path;
                        string relativePath = "../" + Urlconvertor(cutImagePath);
                        //把头像后半部分path保存到DB,你可能不需要
                        //long studentID = -1;
                        //long.TryParse(Session["sid"].ToString(), out studentID);
                        //studentLogic.UpdateAvatarUrl(studentID, finalSmallImgPath);

                        return Json(new { success = true, message = relativePath });
                    }
                    else
                    {
                        return Json(new { success = false, message = "" });
                    }
                }
                catch (Exception)
                {
                    return Json(new { success = false, message = "" });
                }
            }
            catch (Exception)
            {
                return Json(new { success = false, message = "" });
            }
        }
        #endregion

        #region 头像上传辅助函数-相对绝对路径转换函数
        //本地路径转换成URL相对路径
        private string Urlconvertor(string imagesurl1)
        {
            if (System.Web.HttpContext.Current.Request.ApplicationPath != null)
            {
                string tmpRootDir = Server.MapPath(System.Web.HttpContext.Current.Request.ApplicationPath);//获取程序根目录
                string imagesurl2 = imagesurl1.Replace(tmpRootDir, ""); //转换成相对路径
                imagesurl2 = imagesurl2.Replace(@"\", @"/");
                //imagesurl2 = imagesurl2.Replace(@"Aspx_Uc/", @"");
                return imagesurl2;
            }
            return null;
        }
        //相对路径转换成服务器本地物理路径
        public string Urlconvertorlocal(string imagesurl1)
        {
            if (System.Web.HttpContext.Current.Request.ApplicationPath != null)
            {
                string tmpRootDir = Server.MapPath(System.Web.HttpContext.Current.Request.ApplicationPath);//获取程序根目录
                string imagesurl2 = tmpRootDir + imagesurl1.Replace(@"../", @"").Replace(@"./", @"").Replace(@"/", @"\"); //转换成绝对路径
                return imagesurl2;
            }
            return null;
        }
        #endregion

        #endregion

        #region 验证码生成
        [AllowAnonymous]
        public void GetIdentifyImage()
        {
            //context.Response.ContentType = "text/plain";
            //context.Response.Write("Hello World");

            //GDI+绘图技术：三个步骤：1.创建画布 2.画笔，3：为画布绘制所需要的东西

            //隐式类型推断 var 关键字
            string vCode = VerificationCodeHelper.CreateRandomCode(5);
            var imgByte = VerificationCodeHelper.DrawImage(vCode, 20, Color.WhiteSmoke);
            Session["vCode"] = vCode;
            Response.ContentType = "image/gif";
            Response.BinaryWrite(imgByte);
        }
        #endregion

        #region 注销
        public ActionResult LogOut()
        {
            Session.RemoveAll();
            return RedirectToAction("Login");
        }
        #endregion

    }
}