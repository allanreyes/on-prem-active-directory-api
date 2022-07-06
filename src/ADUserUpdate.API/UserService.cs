using Microsoft.Extensions.Options;
using System.DirectoryServices;
using System.DirectoryServices.AccountManagement;

namespace ADUserUpdate.API
{
    internal class UserService : IUserService
    {
        private readonly DirectoryEntry _ldapConnection;

        public UserService(IOptions<AppSettings> appSettings)
        {
            _ldapConnection = new DirectoryEntry(appSettings.Value.Domain)
            {
                Path = appSettings.Value.Path,
                AuthenticationType = AuthenticationTypes.Secure
            };
        }

        public User GetUser(string upn)
        {
            if (GetUserByUPN(upn, out var user))
            {
                return new User()
                {
                    UPN = upn,
                    EmploymentType = user.Properties["employeeType"].Value as string,
                    ManagerUPN = GetManagerUPN(user.Properties["manager"].Value as string),
                    Department = user.Properties["department"].Value as string,
                    Community = user.Properties["l"].Value as string,
                    BuildingName = user.Properties["physicalDeliveryOfficeName"].Value as string,
                    FloorNumber = user.Properties["floorNumber"].Value as string,
                    TelephoneNumber = user.Properties["telephoneNumber"].Value as string,
                    TelephoneNumberExt = user.Properties["telephoneNumberExt"].Value as string,
                    PreferredLanguage = user.Properties["preferredLanguage"].Value as string
                };
            }
            return null;
        }

        private bool GetUserByUPN(string upn, out DirectoryEntry user)
        {
            user = null;

            if (string.IsNullOrWhiteSpace(upn)) return false;

            using var search = new DirectorySearcher(_ldapConnection, $"(&(objectClass=person)(objectClass=user)(userPrincipalName={upn}))");
            search.SearchScope = SearchScope.Subtree;
            var result = search.FindOne();

            if (result == null) return false;

            user = new DirectoryEntry(result.Path);
            return true;
        }

        private string GetManagerUPN(string distinguishedName)
        {
            if (!string.IsNullOrWhiteSpace(distinguishedName))
            {
                using var search = new DirectorySearcher(_ldapConnection, $"(&(objectClass=person)(objectClass=user)(distinguishedName={distinguishedName}))");
                var result = search.FindOne();
                if (result != null)
                {
                    var entryManager = new DirectoryEntry();
                    entryManager = result.GetDirectoryEntry();
                    return entryManager.Properties["userPrincipalName"].Value as string;
                }
            }
            return string.Empty;
        }

        public void UpdateUser(User user)
        {
            if (GetUserByUPN(user.UPN, out var userToUpdate))
            {
                if (string.IsNullOrWhiteSpace(user.EmploymentType))
                    userToUpdate.Properties["employeeType"].Clear();
                else
                    userToUpdate.Properties["employeeType"].Value = user.EmploymentType;


                if (string.IsNullOrEmpty(user.ManagerUPN))
                    userToUpdate.Properties["manager"].Clear();
                else
                    if (GetUserByUPN(user.ManagerUPN, out var manager))
                    userToUpdate.Properties["manager"].Value = manager.Properties["distinguishedName"].Value as string;


                if (string.IsNullOrWhiteSpace(user.Department))
                    userToUpdate.Properties["department"].Clear();
                else
                    userToUpdate.Properties["department"].Value = user.Department;


                userToUpdate.Properties["l"].Value = user.Community;
                if (string.IsNullOrWhiteSpace(user.Department))
                    userToUpdate.Properties["department"].Clear();
                else
                    userToUpdate.Properties["department"].Value = user.Department;


                if (string.IsNullOrWhiteSpace(user.BuildingName))
                    userToUpdate.Properties["physicalDeliveryOfficeName"].Clear();
                else
                    userToUpdate.Properties["physicalDeliveryOfficeName"].Value = user.BuildingName;


                if (string.IsNullOrWhiteSpace(user.FloorNumber))
                    userToUpdate.Properties["floorNumber"].Clear();
                else
                    userToUpdate.Properties["floorNumber"].Value = user.FloorNumber;


                if (string.IsNullOrWhiteSpace(user.TelephoneNumber))
                    userToUpdate.Properties["telephoneNumber"].Clear();
                else
                    userToUpdate.Properties["telephoneNumber"].Value = user.TelephoneNumber;


                if (string.IsNullOrWhiteSpace(user.TelephoneNumberExt))
                    userToUpdate.Properties["telephoneNumberExt"].Clear();
                else
                    userToUpdate.Properties["telephoneNumberExt"].Value = user.TelephoneNumberExt;


                if (string.IsNullOrWhiteSpace(user.PreferredLanguage))
                    userToUpdate.Properties["preferredLanguage"].Clear();
                else
                    userToUpdate.Properties["preferredLanguage"].Value = user.PreferredLanguage;

                userToUpdate.CommitChanges();
            }
        }

        public object GetExpiringUsers(int daysFromToday)
        {
            var filterDate = DateTime.Now.AddDays(daysFromToday);
            var context = new PrincipalContext(ContextType.Domain);
            return UserPrincipal.FindByExpirationTime(context, filterDate, System.DirectoryServices.AccountManagement.MatchType.LessThan)
                .Where(user => user.AccountExpirationDate != null && user.AccountExpirationDate.Value > DateTime.Now)
                .Select(user => new
                {
                    UPN = user.UserPrincipalName,
                    Name = user.Name,
                    AccountExpirationDate = user.AccountExpirationDate.Value.ToLocalTime().ToString("yyyy-MM-dd")
                });
        }

        public object GetExpiredUsers()
        {
            var context = new PrincipalContext(ContextType.Domain);
            return UserPrincipal.FindByExpirationTime(context, DateTime.Now, System.DirectoryServices.AccountManagement.MatchType.LessThan)
                .Where(user => user.AccountExpirationDate != null)
                .Select(user => new
                {
                    UPN = user.UserPrincipalName,
                    Name = user.Name,
                    AccountExpirationDate = user.AccountExpirationDate.Value.ToLocalTime().ToString("yyyy-MM-dd")
                });
        }
    }

    internal interface IUserService
    {
        User GetUser(string upn);

        void UpdateUser(User user);

        object GetExpiringUsers(int daysFromToday);

        object GetExpiredUsers();
    }
}