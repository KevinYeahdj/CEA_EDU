using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using System.Linq;
using System.Text;
using System.Web;
using System.Web.Mvc;
using System.Web.Script.Serialization;
using CEA_EDU.Domain;
using CEA_EDU.Domain.Entity;
using CEA_EDU.Domain.Manager;
using CEA_EDU.Web.Models;
using CEA_EDU.Web.Utils;
using Newtonsoft.Json;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

namespace CEA_EDU.Web.Controllers
{
    public class ManageController : BaseController
    {
        public ActionResult UserIndex()
        {
            DicManager dm = new DicManager();
            var companies = dm.GetDicByType("公司");
            ViewBag.Companies = companies;

            var contents = GetStandardMenuTree(false);
            ViewBag.treeNodes = JsonConvert.SerializeObject(contents);

            return View();
        }
        public ActionResult DicIndex()
        {
            return View();
        }

        public ActionResult PeopleIndex()
        {
            return View();
        }

        public ActionResult MenuIndex()
        {
            var contents = GetStandardMenuTree(true);
            ViewBag.treeNodes = JsonConvert.SerializeObject(contents);
            return View();
        }

        #region 用户方法
        public void GetAllUsers()
        {
            //用于序列化实体类的对象  
            JavaScriptSerializer jss = new JavaScriptSerializer();

            //请求中携带的条件  
            string order = HttpContext.Request.Params["order"];
            string sort = HttpContext.Request.Params["sort"];
            string searchKey = HttpContext.Request.Params["search"];
            int offset = Convert.ToInt32(HttpContext.Request.Params["offset"]);  //0
            int pageSize = Convert.ToInt32(HttpContext.Request.Params["limit"]);

            int total = 0;
            UserManager um = new UserManager();
            List<UserEntity> list = um.GetSearch(searchKey, sort, order, offset, pageSize, out total);

            DicManager dm = new DicManager();
            List<DicEntity> companyDicE = dm.GetDicByType("公司");
            Dictionary<string, string> companyDic = new Dictionary<string, string>();
            foreach (var item in companyDicE)
            {
                companyDic.Add(item.iKey, item.iValue);
            }
            List<UserViewModel> listView = new List<UserViewModel>();
            foreach (var item in list)
            {
                listView.Add(new UserViewModel { iCompanyCode = item.iCompanyCode, iCompanyName = companyDic[item.iCompanyCode], iEmployeeCodeId = item.iEmployeeCodeId, iPassWord = item.iPassWord, iUserName = item.iUserName, iUserType = item.iUserType, iUpdatedOn = item.iUpdatedOn.ToString("yyyyMMdd HH:mm") });
            }

            //给分页实体赋值  
            PageModels<UserViewModel> model = new PageModels<UserViewModel>();
            model.total = total;
            if (total % pageSize == 0)
                model.page = total / pageSize;
            else
                model.page = (total / pageSize) + 1;

            model.rows = listView;

            //将查询结果返回  
            HttpContext.Response.Write(jss.Serialize(model));
        }

        public string UsersSaveChanges(string jsonString, string action)
        {

            try
            {
                UserEntity ue = JsonConvert.DeserializeObject<UserEntity>(jsonString);
                UserManager pm = new UserManager();
                if (action == "add")
                {
                    ue.iCreatedBy = SessionHelper.CurrentUser.iUserName;
                    ue.iUpdatedBy = SessionHelper.CurrentUser.iUserName;
                    pm.Insert(ue);
                }
                else
                {
                    UserEntity ueOld = pm.GetUser(ue.iEmployeeCodeId);
                    ue.iUpdatedBy = SessionHelper.CurrentUser.iUserName;
                    ue.iCreatedBy = ueOld.iCreatedBy;
                    ue.iCreatedOn = ueOld.iCreatedOn;
                    pm.Update(ue);
                }
                return "success";
            }
            catch (Exception e)
            {
                return e.ToString();
            }
        }

        #endregion

        #region 字典方法

        public void GetAllDics()
        {
            //用于序列化实体类的对象  
            JavaScriptSerializer jss = new JavaScriptSerializer();

            //请求中携带的条件  
            string order = HttpContext.Request.Params["order"];
            string sort = HttpContext.Request.Params["sort"];
            string searchKey = HttpContext.Request.Params["search"];
            int offset = Convert.ToInt32(HttpContext.Request.Params["offset"]);  //0
            int pageSize = Convert.ToInt32(HttpContext.Request.Params["limit"]);

            int total = 0;
            DicManager dm = new DicManager();
            List<DicEntity> list = dm.GetSearch(searchKey, sort, order, offset, pageSize, out total);
            List<DicViewModel> listView = new List<DicViewModel>();
            foreach (var item in list)
            {
                listView.Add(new DicViewModel { iId = item.iId, iKey = item.iKey, iValue = item.iValue, iType = item.iType, iUpdatedOn = item.iUpdatedOn.ToString("yyyyMMdd HH:mm") });
            }

            //给分页实体赋值  
            PageModels<DicViewModel> model = new PageModels<DicViewModel>();
            model.total = total;
            if (total % pageSize == 0)
                model.page = total / pageSize;
            else
                model.page = (total / pageSize) + 1;

            model.rows = listView;

            //将查询结果返回  
            HttpContext.Response.Write(jss.Serialize(model));
        }

        public string DicsSaveChanges(string jsonString, string action)
        {

            try
            {
                DicEntity ue = JsonConvert.DeserializeObject<DicEntity>(jsonString);
                DicManager dm = new DicManager();
                if (action == "add")
                {
                    ue.iId = Guid.NewGuid().ToString();
                    ue.iKey = ue.iValue;
                    ue.iCreatedBy = SessionHelper.CurrentUser.iUserName;
                    ue.iUpdatedBy = SessionHelper.CurrentUser.iUserName;
                    dm.Insert(ue);
                }
                else
                {
                    DicEntity ueOld = dm.GetDic(ue.iId);
                    ue.iUpdatedBy = SessionHelper.CurrentUser.iUserName;
                    ue.iKey = ueOld.iKey;
                    ue.iCreatedBy = ueOld.iCreatedBy;
                    ue.iCreatedOn = ueOld.iCreatedOn;
                    dm.Update(ue);
                }
                return "success";
            }
            catch (Exception e)
            {
                return e.ToString();
            }
        }

        #endregion

        #region 公用方法
        public ActionResult ChangeCompany(string newCompany)
        {
            if (SessionHelper.CurrentUser == null)
            {
                return RedirectToAction("Login", "Account");
            }
            else
            {
                UserEntity userinfo = SessionHelper.CurrentUser;
                userinfo.iCompanyCode = newCompany;
                Session[SessionHelper.CurrentUserKey] = userinfo;
                return Content("success");
            }
        }

        private void WriteDataTableToServer(DataTable dt, string targettable, string connectstr)
        {

            using (SqlConnection conn = new SqlConnection(connectstr))
            {
                conn.Open();
                using (SqlBulkCopy sqlbulk = new SqlBulkCopy(conn))
                {

                    sqlbulk.DestinationTableName = targettable;
                    sqlbulk.NotifyAfter = dt.Rows.Count;
                    for (int i = 0; i < dt.Columns.Count; ++i)
                    {
                        SqlBulkCopyColumnMapping map = new SqlBulkCopyColumnMapping(dt.Columns[i].ColumnName, dt.Columns[i].ColumnName);
                        sqlbulk.ColumnMappings.Add(map);

                    }
                    sqlbulk.WriteToServer(dt);

                }

            }

        }
        #endregion


        #region 菜单方法
        public string SaveItemTree(string treeJson)
        {
            try
            {
                List<ItemContent2> contents = JsonConvert.DeserializeObject<List<ItemContent2>>(treeJson);
                List<ItemContent> contents2 = RecursiveFun(contents[0]);
                if (contents2.Count() > 0)
                {
                    string[] nameUrl = contents2[0].name.Split('$');
                    StringBuilder sourceTableSb = new StringBuilder("MERGE INTO [SysMenu] AS T USING( SELECT ");
                    sourceTableSb.Append("'" + contents2[0].id + "' as IGUID, ");
                    sourceTableSb.Append("'" + contents2[0].pId + "' as IPARENTID, ");
                    sourceTableSb.Append("'" + nameUrl[0].Replace("'", "''") + "' as INAME, ");
                    sourceTableSb.Append("'" + nameUrl[1] + "' as IURL ");

                    for (int i = 1; i < contents2.Count; i++)
                    {
                        nameUrl = contents2[i].name.Split('$');
                        sourceTableSb.Append(" union all select ");
                        sourceTableSb.Append("'" + contents2[i].id + "', ");
                        sourceTableSb.Append("'" + contents2[i].pId + "', ");
                        sourceTableSb.Append("'" + nameUrl[0].Replace("'", "''") + "', ");
                        sourceTableSb.Append("'" + nameUrl[1] + "' ");
                    }
                    sourceTableSb.Append(") AS S ON T.IGUID = S.IGUID WHEN MATCHED THEN UPDATE SET T.IPARENTID=S.IPARENTID, T.INAME=S.INAME, T.IUrl = S.IURL WHEN NOT MATCHED THEN INSERT(IGUID,IPARENTID,INAME,IURL) VALUES(S.IGUID,S.IPARENTID,S.INAME,S.IURL);");

                    DbHelperSQL.ExecuteSql(sourceTableSb.ToString());

                }
                /* 样例
       MERGE INTO [TestDb].[dbo].[GUIDTable] AS T
       USING 
     ( select '1' as iguid, 'test1'as idata
     union all select '1375' , 'test1375'
     union all select '2' , 'test2'
     ) 
     AS S
     ON T.iguid=S.iguid
     WHEN MATCHED 
     THEN UPDATE SET T.idata=S.idata , t.iguid=s.iguid
     WHEN NOT MATCHED 
     THEN INSERT(iguid)  VALUES(S.iguid);
     */
                return "success";

            }
            catch (Exception e)
            {
                return "fail" + e.ToString();
            }

        }

        private List<ItemContent> RecursiveFun(ItemContent2 items)
        {
            List<ItemContent> resultTemp = new List<ItemContent>();
            resultTemp.Add(new ItemContent { id = items.id, pId = items.pId, name = items.name });

            if (items.children == null || items.children.Count() == 0)
                return resultTemp;
            else
            {
                foreach (var item in items.children)
                {
                    resultTemp.AddRange(RecursiveFun(item));
                }
                return resultTemp;
            }
        }

        public JsonResult GetUserMenuTree(string userId, string companyId)
        {
            try
            {
                string querySql = "select iMenuId+'|'+iMenuRights from sysUserMenu where iemployeecode = '{0}' and icompanycode= '{1}'";
                DataSet ds = DbHelperSQL.Query(string.Format(querySql, userId, companyId));
                List<string> contents = new List<string>();
                foreach (DataRow dr in ds.Tables[0].Rows)
                {
                    contents.Add(dr[0].ToString());
                }
                return new JsonResult { Data = new { success = true, data = contents }, JsonRequestBehavior = JsonRequestBehavior.AllowGet };
            }
            catch (Exception ex)
            {
                return new JsonResult { Data = new { success = false, data = "", msg = ex.ToString() }, JsonRequestBehavior = JsonRequestBehavior.AllowGet };
            }

        }

        public string UserMenuSaveChanges(string jsonString, string userid, string companyid)
        {
            try
            {
                List<string> menuIds = JsonConvert.DeserializeObject<List<string>>(jsonString);

                DataTable dt = new DataTable("sysUserMenu");
                dt.Columns.Add("iEmployeeCode");
                dt.Columns.Add("iCompanyCode");
                dt.Columns.Add("iMenuId");
                dt.Columns.Add("iMenuRights");
                foreach (string menuid in menuIds)
                {
                    DataRow dataRow = dt.NewRow();
                    dataRow[0] = userid;
                    dataRow[1] = companyid;
                    dataRow[2] = menuid.Split('|')[0];
                    dataRow[3] = menuid.Split('|')[1];
                    dt.Rows.Add(dataRow);
                }
                string deleteSql = "delete from [sysUserMenu] where iemployeecode = '{0}' and icompanycode = '{1}' ";
                DbHelperSQL.ExecuteSql(string.Format(deleteSql, userid, companyid));
                WriteDataTableToServer(dt, "sysUserMenu", ConfigurationManager.ConnectionStrings["HRMSDBConnectionString"].ConnectionString);
                return "success";
            }
            catch (Exception e)
            {
                return e.ToString();
            }
        }

        private List<ItemContent> GetStandardMenuTree(bool widthUrl = true)
        {
            string querySql = "select iguid, iparentid, iname, iUrl from  [sysMenu] ";
            DataSet ds = DbHelperSQL.Query(querySql);
            List<ItemContent> contents = new List<ItemContent>();
            foreach (DataRow dr in ds.Tables[0].Rows)
            {
                if (widthUrl)
                {
                    contents.Add(new ItemContent { id = dr[0].ToString(), pId = dr[1].ToString(), name = dr[2].ToString() + "$" + ((dr[3] != null && !string.IsNullOrEmpty(dr[3].ToString())) ? dr[3].ToString() : ""), open = true });

                }
                else
                {
                    contents.Add(new ItemContent { id = dr[0].ToString(), pId = dr[1].ToString(), name = dr[2].ToString(), open = true });

                }
            }
            return contents;
        }

        public string GetMenuHtml()
        {
            string querySql = "select iguid, iparentid, iname, iUrl from  [sysMenu] a inner join [sysUserMenu] b on a.iguid = b.imenuid and b.iemployeecode='" + SessionHelper.CurrentUser.iEmployeeCodeId + "' and b.icompanycode='" + SessionHelper.CurrentUser.iCompanyCode + "' ";
            if (SessionHelper.CurrentUser.iUserType == "超级管理员")
            {
                querySql = "select iguid, iparentid, iname, iUrl from  [sysMenu]";
            }
            DataSet ds = DbHelperSQL.Query(querySql);
            List<ItemContent> contents = new List<ItemContent>();
            foreach (DataRow dr in ds.Tables[0].Rows)
            {
                contents.Add(new ItemContent { id = dr[0].ToString(), pId = dr[1].ToString(), name = dr[2].ToString() + "$" + ((dr[3] != null && !string.IsNullOrEmpty(dr[3].ToString())) ? dr[3].ToString() : "") });
            }
            StringBuilder sb = new StringBuilder();
            foreach (var item in contents.Where(i => i.pId == "0"))
            {
                sb.AppendLine(string.Format("<li class=\"sub-menu\"> \n <a href=\"javascript:;\"> \n <i class=\"icon-laptop\"></i> \n <span>{0}</span> \n </a>", item.name.Split('$')[0]));

                foreach (var menu in contents.Where(i => i.pId == item.id))
                {
                    sb.AppendLine(string.Format("<ul class=\"sub\"> \n <li><a href=\"{1}\">{0}</a></li> \n </ul>", menu.name.Split('$')[0], menu.name.Split('$')[1]));
                }
                sb.AppendLine("</li>");
            }
            return sb.ToString();
        }
        #endregion
        
    }
}
