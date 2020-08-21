using FolderPermission.Utilities;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using ClosedXML.Excel;

namespace FolderPermission.Controller
{
    class Action
    {
        public static async Task checkAsync(DataTable dt)
        {
            await Task.Run(() =>
            {
                try
                {
                    foreach (DataColumn dc in dt.Columns)
                    {
                        if (dc.ColumnName.ToUpper().Contains(Config.fileServer.ToUpper()))
                        {
                            DataTable userDt = Permission.getPermission(dc.ColumnName);
                            foreach (DataRow dr in userDt.Rows)
                            {
                                string userAccount = dr["UserAccount"].ToString();
                                if (Config.isIgnore(userAccount))
                                {
                                    continue;
                                }
                                var exist = dt.Select("[UserAccount]='" + userAccount + "'");
                                if (exist.Length > 0)
                                {
                                    int index = dt.Rows.IndexOf(exist[0]);
                                    dt.Rows[index][dc.ColumnName] = dr["Permission"];
                                }
                                else
                                {
                                    dt.Rows.Add();
                                    dt.Rows[dt.Rows.Count - 1]["Name"] = dr["Name"];
                                    dt.Rows[dt.Rows.Count - 1]["Role"] = dr["Role"];
                                    dt.Rows[dt.Rows.Count - 1]["E-mail"] = dr["E-mail"];
                                    dt.Rows[dt.Rows.Count - 1]["User ID"] = dr["User ID"];
                                    dt.Rows[dt.Rows.Count - 1]["UserAccount"] = userAccount;
                                    dt.Rows[dt.Rows.Count - 1][dc.ColumnName] = dr["Permission"];
                                }
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.ToString());
                }
            });
        }
        public static async Task setPermissionAsync(DataTable dt)
        {
            await Task.Run(() =>
            {
                try
                {
                    foreach (DataColumn dc in dt.Columns)
                    {
                        if (dc.ColumnName.ToUpper().Contains(Config.fileServer.ToUpper()))
                        {
                            DataTable userDt = Permission.getPermission(dc.ColumnName);
                            // Add, Update
                            foreach (DataRow dr in dt.Rows)
                            {
                                string userAccount = dr["UserAccount"].ToString();
                                var exist = userDt.Select("[UserAccount]='" + userAccount + "'");
                                if (isSamePermissionLevel(exist, dr[dc.ColumnName].ToString()))
                                {
                                    continue; // Same permission
                                }
                                else
                                {
                                    string role = dr["Role"].ToString();
                                    if (role != "-" && role != "")
                                    {
                                        var groupExist = userDt.Select("[Role]='" + role + "'");
                                        bool isGroupExist = Permission.isGroupExist(role);
                                        if (isGroupExist && groupExist.Length > 0)
                                        {
                                            Permission.AddUserToGroup(userAccount, role);
                                        }
                                        else if (isGroupExist)
                                        {
                                            Permission.setPermission(dc.ColumnName, userAccount, dr[dc.ColumnName].ToString());
                                        }
                                        else
                                        {
                                            Permission.createGroup(role);
                                            Permission.AddUserToGroup(userAccount, role);
                                            Permission.setPermission(dc.ColumnName, Environment.MachineName + "\\" + role, dr[dc.ColumnName].ToString());
                                            userDt = Permission.getPermission(dc.ColumnName);
                                        }

                                    }
                                    else
                                    {
                                        Permission.setPermission(dc.ColumnName, userAccount, dr[dc.ColumnName].ToString());
                                    }

                                }
                            }
                            // Remove
                            foreach (DataRow dr in userDt.Rows)
                            {
                                string userAccount = dr["UserAccount"].ToString();
                                var exist = dt.Select("[UserAccount]='" + userAccount + "'");
                                if (Config.isIgnore(userAccount))
                                {
                                    continue;
                                }
                                else if (exist.Length == 0)
                                {
                                    if (dr["Role"].ToString() == "-" || dr["Role"].ToString() == "")
                                    {
                                        Permission.removePermission(dc.ColumnName, userAccount);
                                    }
                                    else
                                    {
                                        Permission.RemoveUserFromGroup(userAccount, dr["Role"].ToString());
                                    }
                                }
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.ToString());
                }
            });
        }
        public static async Task auditAsync(DataTable dt, string passString, string failedString)
        {
            await Task.Run(() =>
            {
                try
                {
                    foreach (DataColumn dc in dt.Columns)
                    {
                        if (dc.ColumnName.ToUpper().Contains(Config.fileServer.ToUpper()))
                        {
                            DataTable userDt = Permission.getPermission(dc.ColumnName);
                            // Add, Update
                            foreach (DataRow dr in dt.Rows)
                            {
                                string userAccount = dr["UserAccount"].ToString();
                                var exist = userDt.Select("[UserAccount]='" + userAccount + "'");
                                if (isSamePermissionLevel(exist, dr[dc.ColumnName].ToString()))
                                {
                                    dr[dc.ColumnName] = dr[dc.ColumnName].ToString() + " [" + passString + "]";
                                }
                                else if (exist.Length > 0 && exist[0]["Permission"].ToString() == "")
                                {
                                    continue;
                                }
                                else
                                {
                                    dr[dc.ColumnName] = dr[dc.ColumnName].ToString() + " [" + failedString + "]";
                                }
                            }
                            // Remove
                            foreach (DataRow dr in userDt.Rows)
                            {
                                string userAccount = dr["UserAccount"].ToString();
                                var exist = dt.Select("[UserAccount]='" + userAccount + "'");
                                if (Config.isIgnore(userAccount))
                                {
                                    continue;
                                }
                                else if (exist.Length == 0)
                                {
                                    dt.Rows.Add();
                                    dt.Rows[dt.Rows.Count - 1]["Name"] = dr["Name"];
                                    dt.Rows[dt.Rows.Count - 1]["Role"] = dr["Role"];
                                    dt.Rows[dt.Rows.Count - 1]["E-mail"] = dr["E-mail"];
                                    dt.Rows[dt.Rows.Count - 1]["User ID"] = dr["User ID"];
                                    dt.Rows[dt.Rows.Count - 1]["UserAccount"] = dr["UserAccount"];
                                    dt.Rows[dt.Rows.Count - 1][dc.ColumnName] = dr["Permission"] + " [" + failedString + "]";
                                }
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.ToString());
                }
            });
        }
        public static void exportToExcel(string outputFile, DataTable dt)
        {
            using (var workbook = new XLWorkbook())
            {
                var worksheet = workbook.Worksheets.Add(dt,"Permission");
                int i = 1;
                foreach (DataColumn dc in dt.Columns)
                {
                    worksheet.Cell(1, i).Value = dc.ColumnName;
                    i++;
                }
                worksheet.Columns(1, i).AdjustToContents();
                workbook.SaveAs(outputFile);
            }
        }
        private static bool isSamePermissionLevel(DataRow[] rows, string permission)
        {
            bool isSame = false;
            foreach (DataRow dr in rows)
            {
                if (dr["Permission"].ToString() == permission)
                {
                    isSame = true;
                    break;
                }
            }

            return isSame;
        }
        
    }
}
