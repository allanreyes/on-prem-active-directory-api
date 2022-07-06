using Microsoft.Extensions.Options;
using System.DirectoryServices;
using System.DirectoryServices.AccountManagement;

namespace ADUserUpdate.API
{
    internal class UserService : IUserService
    {
        private readonly DirectoryEntry _ldapConnection;
        private readonly PrincipalContext _context;

        public UserService(IOptions<AppSettings> appSettings)
        {
            _ldapConnection = new DirectoryEntry(appSettings.Value.Domain)
            {
                Path = appSettings.Value.Path,
                AuthenticationType = AuthenticationTypes.Secure
            };
            _context = new PrincipalContext(ContextType.Domain);
        }

        public User GetUser(string upn)
        {
            if (GetUserByUPN(upn, out var user))
            {
                return new User()
                {
                    UPN = upn,
                    EmploymentType = user.Properties[Constants.EmploymentType].Value as string,
                    ManagerUPN = GetManagerUPN(user.Properties[Constants.Manager].Value as string),
                    Department = user.Properties[Constants.Department].Value as string,
                    Community = user.Properties[Constants.Community].Value as string,
                    BuildingName = user.Properties[Constants.BuildingName].Value as string,
                    FloorNumber = user.Properties[Constants.FloorNumber].Value as string,
                    TelephoneNumber = user.Properties[Constants.TelephoneNumber].Value as string,
                    TelephoneNumberExt = user.Properties[Constants.TelephoneNumberExt].Value as string,
                    PreferredLanguage = user.Properties[Constants.PreferredLanguage].Value as string
                };
            }
            return null;
        }

        private bool GetUserByUPN(string upn, out DirectoryEntry user)
        {
            user = null;

            if (string.IsNullOrWhiteSpace(upn)) return false;

            var p = UserPrincipal.FindByIdentity(_context, IdentityType.UserPrincipalName, upn);
            if (p == null) return false;

            user = p.GetUnderlyingObject() as DirectoryEntry;
            return true;
        }

        private string GetManagerUPN(string distinguishedName)
        {
            if (!string.IsNullOrWhiteSpace(distinguishedName))
            {
                var manager = UserPrincipal.FindByIdentity(_context, IdentityType.DistinguishedName, distinguishedName);
                if (manager != null)
                    return manager.UserPrincipalName;
            }
            return string.Empty;
        }

        public void UpdateUser(User user)
        {
            if (GetUserByUPN(user.UPN, out var userToUpdate))
            {
                if (string.IsNullOrWhiteSpace(user.EmploymentType))
                    userToUpdate.Properties[Constants.EmploymentType].Clear();
                else
                    userToUpdate.Properties[Constants.EmploymentType].Value = user.EmploymentType;


                if (string.IsNullOrEmpty(user.ManagerUPN))
                    userToUpdate.Properties[Constants.Manager].Clear();
                else
                    if (GetUserByUPN(user.ManagerUPN, out var manager))
                    userToUpdate.Properties[Constants.Manager].Value = manager.Properties[Constants.DistinguishedName].Value as string;


                if (string.IsNullOrWhiteSpace(user.Department))
                    userToUpdate.Properties[Constants.Department].Clear();
                else
                    userToUpdate.Properties[Constants.Department].Value = user.Department;


                if (string.IsNullOrWhiteSpace(user.Community))
                    userToUpdate.Properties[Constants.Community].Clear();
                else
                    userToUpdate.Properties[Constants.Community].Value = user.Community;


                if (string.IsNullOrWhiteSpace(user.BuildingName))
                    userToUpdate.Properties[Constants.BuildingName].Clear();
                else
                    userToUpdate.Properties[Constants.BuildingName].Value = user.BuildingName;


                if (string.IsNullOrWhiteSpace(user.FloorNumber))
                    userToUpdate.Properties[Constants.FloorNumber].Clear();
                else
                    userToUpdate.Properties[Constants.FloorNumber].Value = user.FloorNumber;


                if (string.IsNullOrWhiteSpace(user.TelephoneNumber))
                    userToUpdate.Properties[Constants.TelephoneNumber].Clear();
                else
                    userToUpdate.Properties[Constants.TelephoneNumber].Value = user.TelephoneNumber;


                if (string.IsNullOrWhiteSpace(user.TelephoneNumberExt))
                    userToUpdate.Properties[Constants.TelephoneNumberExt].Clear();
                else
                    userToUpdate.Properties[Constants.TelephoneNumberExt].Value = user.TelephoneNumberExt;


                if (string.IsNullOrWhiteSpace(user.PreferredLanguage))
                    userToUpdate.Properties[Constants.PreferredLanguage].Clear();
                else
                    userToUpdate.Properties[Constants.PreferredLanguage].Value = user.PreferredLanguage;

                userToUpdate.CommitChanges();
            }
        }

        public object GetExpiringUsers(int daysFromToday)
        {
            var filterDate = DateTime.Now.AddDays(daysFromToday);
            return UserPrincipal.FindByExpirationTime(_context, filterDate, System.DirectoryServices.AccountManagement.MatchType.LessThan)
                .Where(user => user.AccountExpirationDate != null && user.AccountExpirationDate.Value > DateTime.Now)
                .Select(user => new
                {
                    UPN = user.UserPrincipalName,
                    Name = user.Name,
                    GivenName = user.GivenName,
                    AccountExpirationDate = user.AccountExpirationDate.Value.ToLocalTime().ToString(Constants.DateFormat),
                    Enabled = user.Enabled.Value
                });
        }

        public object GetExpiredUsers()
        {
            return UserPrincipal.FindByExpirationTime(_context, DateTime.Now, System.DirectoryServices.AccountManagement.MatchType.LessThan)
                .Where(user => user.AccountExpirationDate != null)
                .Select(user => new
                {
                    UPN = user.UserPrincipalName,
                    Name = user.Name,
                    GivenName = user.GivenName,
                    AccountExpirationDate = user.AccountExpirationDate.Value.ToLocalTime().ToString(Constants.DateFormat),
                    Enabled = user.Enabled.Value
                });
        }

        public void DisableUser(string upn)
        {
            if (string.IsNullOrWhiteSpace(upn)) return;

            var p = UserPrincipal.FindByIdentity(_context, IdentityType.UserPrincipalName, upn);
            if (p == null) return;
            p.Enabled = false;
            p.Save(_context);
        }
    }

    internal interface IUserService
    {
        User GetUser(string upn);

        void UpdateUser(User user);

        object GetExpiringUsers(int daysFromToday);

        object GetExpiredUsers();

        void DisableUser(string upn);
    }
}