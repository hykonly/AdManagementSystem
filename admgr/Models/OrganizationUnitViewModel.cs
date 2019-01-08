using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace admgr.Models
{
    public class OrganizationUnitViewModel
    {
        public string OrganizationUnitId { get; set; }
        public string OrganizationUnitName { get; set; }
    }

    public class CompanyAndDeptViewModel
    {
        public List<CompanyViewModel> CompanyViewModelList { get; set; }
        public List<DeptViewModel> DeptViewModelList { get; set; }

    }
    public class CompanyContainDeptViewModel
    {
        public string CompanyId { get; set; }
        public List<DeptViewModel> DeptViewModelList { get; set; }
    }
    public class CompanyViewModel
    {
        public string CompanyId { get; set; }
        public string CompanyName { get; set; }
    }
    public class CompanyViewModelTest
    {
        public int CompanyId { get; set; }
        public string CompanyName { get; set; }
    }
    public class DeptViewModel
    {
        public string CompanyId { get; set; }
        public string DeptId { get; set; }
        public string DeptName { get; set; }
    }
    public class RoleViewModel
    {
        public string RoleId { get; set; }
        public string RoleName { get; set; }
    }
    public class EmailDataViewModel
    {
        public int EmailDataId { get; set; }
        public string CompanyId { get; set; }
        public string MainMailAddress { get; set; }
        public string SecondMailAddress { set; get; }
        public string EmailStoreDb { get; set; }
    }
}