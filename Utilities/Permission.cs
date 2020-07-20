using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics.Eventing.Reader;
using System.DirectoryServices;
using System.DirectoryServices.AccountManagement;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Security.AccessControl;
using System.Security.Principal;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace FolderPermission.Utilities
{
    class Permission
    {
        private static readonly FileSystemRights READ = FileSystemRights.Read | FileSystemRights.ReadAndExecute | FileSystemRights.ListDirectory | FileSystemRights.ReadAttributes | FileSystemRights.ReadPermissions;
        private static readonly FileSystemRights WRITE = READ | FileSystemRights.Write | FileSystemRights.Modify;
        public static DataTable getPermission(string dir)
        {
            DataTable dt = new DataTable();
            dt.Columns.Add("Name", typeof(string));
            dt.Columns.Add("Role", typeof(string));
            dt.Columns.Add("E-mail", typeof(string));
            dt.Columns.Add("User ID", typeof(string));
            dt.Columns.Add("UserAccount", typeof(string));
            dt.Columns.Add("Permission", typeof(string));
            DirectoryInfo di = new DirectoryInfo(dir);
            DirectorySecurity acl = di.GetAccessControl(AccessControlSections.Access);
            AuthorizationRuleCollection rules = acl.GetAccessRules(true, true, typeof(NTAccount));
            foreach (AuthorizationRule rule in rules)
            {
                if (Config.isIgnore(rule.IdentityReference.Value))
                {
                    continue;
                }
                var filesystemAccessRule = (FileSystemAccessRule)rule;
                string permission = "";
                if ((filesystemAccessRule.FileSystemRights & FileSystemRights.WriteData) > 0 && filesystemAccessRule.AccessControlType != AccessControlType.Deny)
                {
                    permission = "R/W";
                } else if ((filesystemAccessRule.FileSystemRights & FileSystemRights.ReadData) > 0 && filesystemAccessRule.AccessControlType != AccessControlType.Deny)
                {
                    permission = "R";
                }
                
                if (rule.IdentityReference.Value.ToUpper().Contains(Config.fileServer.ToUpper()))
                {
                    
                    // Groups
                    if (isGroupExist(rule.IdentityReference.Value.ToLower())) {
                        var memberList = getMembers(rule.IdentityReference.Value);
                        foreach (var member in memberList)
                        {
                            string group = rule.IdentityReference.Value.Replace(Config.fileServer + "\\", "");
                            string strIdx = ",DC=";
                            int pFrom = member.DistinguishedName.IndexOf(strIdx) + strIdx.Length;
                            string domain = member.DistinguishedName.Substring(pFrom);
                            int pTo = domain.LastIndexOf(strIdx);
                            domain = domain.Substring(0, pTo);
                            dt.Rows.Add(member.DisplayName, group, member.EmailAddress, member.EmployeeId, domain + "\\" + member.SamAccountName, permission);
                        }
                    } else
                    {
                        dt.Rows.Add("-", "-", "-", "-", rule.IdentityReference.Value, permission);
                    }
                    
                }
                else
                {
                    // Individual User
                    string domain = rule.IdentityReference.Value.Split('\\')[0] + Config.domain;
                    UserPrincipal member = getMemberInfo(rule.IdentityReference.Value, domain);
                    if (member != null)
                    {
                        dt.Rows.Add(member.DisplayName, "-", member.EmailAddress, member.EmployeeId, rule.IdentityReference.Value, permission);
                    } else
                    {
                        dt.Rows.Add("-", "-", "-", "-", rule.IdentityReference.Value, permission);
                    }
                    
                }
            }
            return dt;
        }
        public static void setPermission(string dir, string userAccount, string permission)
        {
            FileSystemRights right = 0;
            if (permission == "R/W")
            {
                right = WRITE;
            }
            else if (permission == "R")
            {
                right = READ;
            }
            if (right != 0)
            {
                SetPermissionDirectory(dir, @userAccount,  right);
            }
            
        }
        public static void removePermission(string dir, string userAccount)
        {

            DirectoryInfo dirinfo = new DirectoryInfo(dir);
            DirectorySecurity dsec = dirinfo.GetAccessControl(AccessControlSections.Access);

            AuthorizationRuleCollection rules = dsec.GetAccessRules(true, true, typeof(System.Security.Principal.NTAccount));
            foreach (AccessRule rule in rules)
            {
                if (rule.IdentityReference.Value == userAccount)
                {
                    bool value;
                    dsec.PurgeAccessRules(rule.IdentityReference);
                    dsec.ModifyAccessRule(AccessControlModification.RemoveAll, rule, out value);
                    dirinfo.SetAccessControl(dsec);
                }
            }
        }
        public static void searchDir(string path, List<string> output, decimal deepLimit)
        {
            if (deepLimit > 0)
            {
                string[] dirs = Directory.GetDirectories(path);
                if (dirs.Length == 0)
                {
                    output.Add(path);
                } else
                {
                    output.AddRange(dirs);
                    foreach (string dir in dirs)
                    {
                        searchDir(dir, output, --deepLimit);
                    }
                }
            }
        }
        public static void SetPermissionDirectory(string DirectoryName, string UserAccount, FileSystemRights right)
        {
            if (!Directory.Exists(DirectoryName))
            {
                Directory.CreateDirectory(DirectoryName);
            }
            
             AddUsersAndPermissions(DirectoryName, UserAccount, right, AccessControlType.Allow);
            
        }

        public static void AddUsersAndPermissions(string DirectoryName, string UserAccount, FileSystemRights UserRights, AccessControlType AccessType)
        {
            // Create a DirectoryInfo object.
            DirectoryInfo directoryInfo = new DirectoryInfo(DirectoryName);

            // Get security settings.
            DirectorySecurity dirSecurity = directoryInfo.GetAccessControl();

            // Add the FileSystemAccessRule to the security settings.
            dirSecurity.RemoveAccessRule(new FileSystemAccessRule(UserAccount, FileSystemRights.FullControl, InheritanceFlags.ContainerInherit | InheritanceFlags.ObjectInherit, PropagationFlags.None, AccessType));
            dirSecurity.AddAccessRule(new FileSystemAccessRule(UserAccount, UserRights, InheritanceFlags.ContainerInherit | InheritanceFlags.ObjectInherit, PropagationFlags.None, AccessType));
            
            // Set the access settings.
            directoryInfo.SetAccessControl(dirSecurity);
        }
        public static List<UserPrincipal> getMembers(string groupName)
        {
            
            List<UserPrincipal> result = new List<UserPrincipal>();

            try
            {
                // establish domain context
                PrincipalContext yourDomain = new PrincipalContext(ContextType.Machine);

                // find your user
                GroupPrincipal group = GroupPrincipal.FindByIdentity(yourDomain, groupName);
                // if found - grab its groups
                if (group != null)
                {
                    // iterate over all groups
                    foreach (Principal p in group.GetMembers())
                    {
                        UserPrincipal theUser = p as UserPrincipal;
                        if (theUser != null)
                        {
                            result.Add(theUser);
                        }
                    }
                    group.Dispose();
                }
            } catch (Exception)
            {
                result = new List<UserPrincipal>();
                string name = groupName.Replace(Environment.MachineName + "\\", "");

                string directoryString = "WinNT://./" + name + ",group";
                using (DirectoryEntry groupEntry = new DirectoryEntry(directoryString))
                {
                    foreach (object member in (IEnumerable)groupEntry.Invoke("Members"))
                    {
                        using (DirectoryEntry memberEntry = new DirectoryEntry(member))
                        {
                            string userIdValue = memberEntry.Path;
                            //output will be WinNT://system1/user1
                            //split with '//' to the user code
                            string[] user = userIdValue.Split(new char[] { '\\' });
                            string userId = user[user.Length - 1].Replace("WinNT://", "").Replace("/",@"\");
                            string domain = userId.Split('\\')[0] + Config.domain;

                            result.Add(getMemberInfo(userId, domain));

                        }
                    }
                }

            }
            
            

            return result;
        }
        public static UserPrincipal getMemberInfo(string name, string domain="")
        {
            try
            {
                PrincipalContext ctx = new PrincipalContext(ContextType.Domain);
                if (domain != "")
                {
                    ctx = new PrincipalContext(ContextType.Domain, domain);
                }
                UserPrincipal user = UserPrincipal.FindByIdentity(ctx, name);

                if (user != null)
                {
                    return user;
                }
            }
            catch (Exception)
            {

            }
            

            return null;
        }
        public static void createGroup( string groupName)
        {
            try
            {
                using (PrincipalContext pc = new PrincipalContext(ContextType.Machine))
                {
                    GroupPrincipal group = new GroupPrincipal(pc);
                    group.Name = groupName;
                    group.Save();
                }
            }
            catch (Exception ex)
            {
                //doSomething with E.Message.ToString(); 
                MessageBox.Show(ex.ToString());
            }
        }
        public static bool isGroupExist(string groupName)
        {
            bool isExist = false;
            using (PrincipalContext pc = new PrincipalContext(ContextType.Machine))
            {
                try
                {
                    GroupPrincipal group = GroupPrincipal.FindByIdentity(pc, groupName);
                    if (group != null)
                    {
                        isExist = true;
                    }
                }
                catch (Exception)
                {

                }
                
            }
            return isExist;

        }
        public static void AddUserToGroup(string userId, string groupName)
        {
            try
            {
                using (PrincipalContext pc = new PrincipalContext(ContextType.Machine))
                {
                    GroupPrincipal group = GroupPrincipal.FindByIdentity(pc, groupName);
                    group.Members.Add(getMemberInfo(userId));
                    group.Save();
                }
            }
            catch (Exception)
            {
                //doSomething with E.Message.ToString(); 
                string[] splited = userId.Split('\\');
                string domain = splited[0];
                string userPath = string.Format("WinNT://{0}/{1},user", domain, splited[1]);
                string groupPath = string.Format("WinNT://{0}/{1},group", Environment.MachineName, groupName);
                using (DirectoryEntry group = new DirectoryEntry(groupPath))
                {
                    group.Invoke("Add", userPath);
                    group.CommitChanges();
                }
            }
        }

        public static void RemoveUserFromGroup(string userId, string groupName)
        {
            try
            {
                using (PrincipalContext pc = new PrincipalContext(ContextType.Machine))
                {
                    GroupPrincipal group = GroupPrincipal.FindByIdentity(pc, groupName);
                    group.Members.Remove(getMemberInfo(userId));
                    group.Save();
                }
            }
            catch (Exception)
            {
                string[] splited = userId.Split('\\');
                string domain = splited[0];
                string userPath = string.Format("WinNT://{0}/{1},user", domain, splited[1]);
                string groupPath = string.Format("WinNT://{0}/{1},group", Environment.MachineName, groupName);
                using (DirectoryEntry group = new DirectoryEntry(groupPath))
                {
                    group.Invoke("Remove", userPath);
                    group.CommitChanges();
                }
            }
        }
    }
}
