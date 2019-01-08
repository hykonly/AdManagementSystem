using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace admgr.Controllers
{
    public class ExchangeController
    {
        //#region 读取exchange中公司数据库对应表
        //[CustomAuthorize(Roles = "44")]
        //public void ReadExchangeDatabase()
        //{
        //    ViewData["role"] = Session["RoleType"].ToString();
        //    AdMgrExchange adMgrExchange = (AdMgrExchange)Session["adMgrExchange"];
        //    List<MailEntity> mailEntities = adMgrExchange.GetEmailsEntitiesy();
        //    var datajson = JsonConvert.SerializeObject(mailEntities);
        //    string callbackFunName = System.Web.HttpContext.Current.Request["callback"];
        //    System.Web.HttpContext.Current.Response.Write(callbackFunName + "(" + datajson + ")");
        //}
        //#endregion

        //#region 更新exchange中公司数据库对应表中数据
        //[CustomAuthorize(Roles = "44")]
        //public void UpdateExchangeDatabase()
        //{
        //    ViewData["role"] = Session["RoleType"].ToString();
        //    string models = System.Web.HttpContext.Current.Request["models"];
        //    List<MailEntity> mailEntities = JsonConvert.DeserializeObject<List<MailEntity>>(models);
        //    string id = mailEntities[0].Id;
        //    string company = mailEntities[0].Company;
        //    string adName = mailEntities[0].AdName;
        //    string emailsData = mailEntities[0].EmailsData;
        //    AdMgrExchange adMgrExchange = GenerateAdMgrExchange();
        //    adMgrExchange.UpdataEmailsData(company, adName, emailsData, id);
        //    MailEntity mailEntity = adMgrExchange.GetEmailsEntityById(id);
        //    var datajson = JsonConvert.SerializeObject(mailEntity);
        //    string callbackFunName = System.Web.HttpContext.Current.Request["callback"];
        //    System.Web.HttpContext.Current.Response.Write(callbackFunName + "(" + datajson + ")");
        //}
        //#endregion

        //#region 删除exchange中公司数据库对应表中数据
        //[CustomAuthorize(Roles = "44")]
        //public void DeleteExchangeDatabase()
        //{
        //    ViewData["role"] = Session["RoleType"].ToString();
        //    string models = System.Web.HttpContext.Current.Request["models"];
        //    List<MailEntity> mailEntities = JsonConvert.DeserializeObject<List<MailEntity>>(models);
        //    string id = mailEntities[0].Id;
        //    string company = mailEntities[0].Company;
        //    string adName = mailEntities[0].AdName;
        //    string emailsData = mailEntities[0].EmailsData;
        //    AdMgrExchange adMgrExchange = GenerateAdMgrExchange();
        //    int i = adMgrExchange.DeleteEmailsData(id, company, adName, emailsData);
        //    //下面的返回被删除的数据，因为这是kendo插件的需要
        //    var datajson = JsonConvert.SerializeObject(mailEntities[0]);
        //    string callbackFunName = System.Web.HttpContext.Current.Request["callback"];
        //    System.Web.HttpContext.Current.Response.Write(callbackFunName + "(" + datajson + ")");

        //}
        //#endregion

        //#region 插入exchange中公司数据库对应表中数据
        //[CustomAuthorize(Roles = "44")]
        //public void InsertExchangeDatabase()
        //{
        //    ViewData["role"] = Session["RoleType"].ToString();
        //    string models = System.Web.HttpContext.Current.Request["models"];
        //    List<MailEntity> mailEntities = JsonConvert.DeserializeObject<List<MailEntity>>(models);
        //    string company = mailEntities[0].Company;
        //    string adName = mailEntities[0].AdName;
        //    string emailsData = mailEntities[0].EmailsData;
        //    AdMgrExchange adMgrExchange = GenerateAdMgrExchange();
        //    string id = adMgrExchange.InsertEmailsDataSingle(company, adName, emailsData);
        //    if (id != null)
        //    {
        //        MailEntity mailEntity = adMgrExchange.GetEmailsEntityById(id);
        //        var datajson = JsonConvert.SerializeObject(mailEntity);
        //        string callbackFunName = System.Web.HttpContext.Current.Request["callback"];
        //        System.Web.HttpContext.Current.Response.Write(callbackFunName + "(" + datajson + ")");
        //    }

        //}
        //#endregion

        //#region Exchange邮箱系统配置
        ///// <summary>
        ///// Exchange邮箱系统配置，主要配置公司名称域邮箱数据库的对应关系
        ///// </summary>
        ///// <returns></returns>
        //[CustomAuthorize(Roles = "44")]
        //public ActionResult ExchangeConfig()
        //{
        //    ViewData["role"] = Session["RoleType"].ToString();
        //    AdMgrExchange adMgrExchange = GenerateAdMgrExchange();
        //    MainSecondMail mainSecondMail = adMgrExchange.GetMainSecondMail();
        //    MainSecondMailModel mainSecondMailModel = new MainSecondMailModel(mainSecondMail);
        //    return View(mainSecondMailModel);
        //}
        //#endregion

        //#region 配置主邮箱、辅邮箱地址
        //[CustomAuthorize(Roles = "44")]
        //[HttpPost]
        //public ActionResult ExchangeConfig(MainSecondMailModel mainSecondMailModel)
        //{
        //    string mainMailAddress = mainSecondMailModel.MainMailAddress;
        //    string secondMailAddress = mainSecondMailModel.SecondMailAddress;
        //    if (mainMailAddress.Contains("@") && mainMailAddress.Contains(".") && secondMailAddress.Contains("@") && secondMailAddress.Contains("."))
        //    {
        //        AdMgrExchange adMgrExchange = GenerateAdMgrExchange();
        //        bool result = adMgrExchange.UpdateMainSecondMail(mainMailAddress, secondMailAddress);
        //        if (result)
        //        {
        //            return Json(true, JsonRequestBehavior.AllowGet);
        //        }

        //    }
        //    return Json(false, JsonRequestBehavior.AllowGet);
        //}
        //#endregion
    }
}