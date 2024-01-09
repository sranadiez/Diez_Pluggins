using System;
using System.Collections.Generic;
using System.Linq;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;

namespace IFZA_application_license
{
    public class IfzaLicenseApplicationPlugin : IPlugin
    {
        private ITracingService tracingService;

        public void Execute(IServiceProvider serviceProvider)
        {
            tracingService = (ITracingService)serviceProvider.GetService(typeof(ITracingService));
            tracingService.Trace("line1");

            try
            {
                IPluginExecutionContext context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
                tracingService.Trace("line2");

                IOrganizationServiceFactory serviceFactory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
                IOrganizationService service = serviceFactory.CreateOrganizationService(context.UserId);

                if (context.InputParameters.Contains("Target") && context.InputParameters["Target"] is Entity)
                {
                    Entity ifzaLicenseApplication = (Entity)context.InputParameters["Target"];
                    Entity PREifzaLicenseApplication = context.PreEntityImages.TryGetValue("preImage", out var preImage)
                        ? preImage
                        : null;
                    tracingService.Trace("line3");

                    if (ifzaLicenseApplication.LogicalName == "ntw_ifzalicenseapplication")
                    {
                        tracingService.Trace("line32");

                        if (PREifzaLicenseApplication != null && PREifzaLicenseApplication.Contains("ntw_companytype") && PREifzaLicenseApplication.Contains("ntw_submitforapproval"))
                        {
                            tracingService.Trace("line35");

                            int companyType = PREifzaLicenseApplication.GetAttributeValue<OptionSetValue>("ntw_companytype")?.Value ?? 0;
                            Guid licenseApplication = PREifzaLicenseApplication.GetAttributeValue<EntityReference>("ntw_licenseapplicationid")?.Id ?? Guid.Empty;
                            tracingService.Trace("licenseApplication: "+ licenseApplication.ToString());
                            bool ds = PREifzaLicenseApplication.GetAttributeValue<bool>("ntw_submitforapproval");
                            tracingService.Trace("line4");

                            if (ds)
                            {
                                switch (companyType)
                                {
                                    case 961110002: // FZCO
                                        ValidateFZCO(ifzaLicenseApplication, service, licenseApplication);
                                        tracingService.Trace("line5");
                                        break;

                                    case 860920000: // PLC
                                        ValidatePLC(ifzaLicenseApplication, service, licenseApplication);
                                        tracingService.Trace("line6");
                                        break;

                                    case 961110000: // Branch
                                        ValidateBranch(ifzaLicenseApplication, service, licenseApplication);
                                        tracingService.Trace("line7");
                                        break;

                                    default:
                                        // Handle other cases if needed
                                        break;
                                }
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                tracingService.Trace($"Error: {ex.Message}");
                throw;
            }
        }

        private void ValidateFZCO(Entity ifzaLicenseApplication, IOrganizationService service, Guid licenseApplication)
        {
            int requiredDirectorCount = 1;
            int requiredManagerCount = 1;
            int requiredCompanySecretaryCount = 1;

            EntityCollection members = FetchMembers(ifzaLicenseApplication, service, licenseApplication);

            int directorCount = CountMembersWithRole(members, "961110001"); //Director
            int managerCount = CountMembersWithRole(members, "961110000"); //Manager
            int companySecretaryCount = CountMembersWithRole(members, "961110003"); //Company Secretary

            // Check if individual roles meet the minimum requirements
            if (directorCount < requiredDirectorCount ||
                managerCount < requiredManagerCount ||
                companySecretaryCount < requiredCompanySecretaryCount)
            {
                throw new InvalidPluginExecutionException("FZCO does not meet the minimum role requirements.");
            }

            ValidateContactProperties(members, service);
        }

        private void ValidatePLC(Entity ifzaLicenseApplication, IOrganizationService service, Guid licenseApplication)
        {
            int requiredDirectorCount = 2;
            int requiredManagerCount = 1;
            int requiredCompanySecretaryCount = 1;

            EntityCollection members = FetchMembers(ifzaLicenseApplication, service, licenseApplication);
            

            int directorCount = CountMembersWithRole(members, "961110001"); //Director
            int managerCount = CountMembersWithRole(members, "961110000"); //Manager
            int companySecretaryCount = CountMembersWithRole(members, "961110003"); //Company Secretary

            // Check if individual roles meet the minimum requirements
            if (directorCount < requiredDirectorCount ||
                managerCount < requiredManagerCount ||
                companySecretaryCount < requiredCompanySecretaryCount)
            {
                throw new InvalidPluginExecutionException("PLC does not meet the minimum role requirements.");
            }

            // Check for the assignment of Director and Company Secretary roles
            CheckDirectorAndCompanySecretaryAssignment(members, service);

            ValidateContactProperties(members, service);
        }

        private void ValidateBranch(Entity ifzaLicenseApplication, IOrganizationService service, Guid licenseApplication)
        {
            int requiredManagerCount = 1;

            EntityCollection members = FetchMembers(ifzaLicenseApplication, service, licenseApplication);
            if(members.Entities.Count>0)
            {
               
                int managerCount = CountMembersWithRole(members, "961110000"); //Manager

                if (managerCount < requiredManagerCount)
                {
                    throw new InvalidPluginExecutionException("Branch requires at least 1 Manager.");
                }

            }
            else
                throw new InvalidPluginExecutionException("No Member Found.");

          
            ValidateContactProperties(members, service);
        }

        private void CheckDirectorAndCompanySecretaryAssignment(EntityCollection members, IOrganizationService service)
        {
            // Replace YourDirectorRoleCode and YourSecretaryRoleCode with actual role codes
            int YourDirectorRoleCode = 961110001; // Replace with the actual value for "Director" in the multi-option set
            int YourSecretaryRoleCode = 961110003; // Replace with the actual value for "Company Secretary" in the multi-option set

            // Get the contact IDs for Directors
            List<Guid> directorIds = members.Entities
    .Where(member => RoleExists(member.GetAttributeValue<OptionSetValueCollection>("ntw_roles"), YourDirectorRoleCode))
    .Select(member => member.GetAttributeValue<EntityReference>("ntw_contact")?.Id ?? Guid.Empty)
    .Where(contactId => contactId != Guid.Empty)
    .ToList();

            // Get the contact IDs for Company Secretaries
            List<Guid> secretaryIds = members.Entities
    .Where(member => RoleExists(member.GetAttributeValue<OptionSetValueCollection>("ntw_roles"), YourSecretaryRoleCode))
    .Select(member => member.GetAttributeValue<EntityReference>("ntw_contact")?.Id ?? Guid.Empty)
    .Where(contactId => contactId != Guid.Empty)
    .ToList();

    
            // Find common contacts between Directors and Company Secretaries using FetchXML
            // Construct the FetchXML to retrieve contacts with specific roles
            string fetchXml = $@"
<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
    <entity name='ntw_licensecontactmapping'>
        <attribute name='ntw_contact' />
        <attribute name='ntw_roles' />
        <filter type='and'>
            <condition attribute='ntw_contact' operator='in'>
                {string.Join("", directorIds.Concat(secretaryIds).Select(id => $"<value>{id}</value>"))}
            </condition>
        </filter>
    </entity>
</fetch>";





            EntityCollection commonContacts = service.RetrieveMultiple(new FetchExpression(fetchXml));

            if (commonContacts.Entities.Count > 0)
            {
                throw new InvalidPluginExecutionException("PLC cannot have the same contact assigned as both Director and Company Secretary.");
            }
        }

        // Function to check if a role exists in the OptionSetValueCollection
        private bool RoleExists(OptionSetValueCollection optionSetValues, int roleCode)
        {
            return optionSetValues != null && optionSetValues.Any(opt => opt.Value == roleCode);
        }


        private void ValidateContactProperties(EntityCollection members, IOrganizationService service)
        {
            foreach (var member in members.Entities)
            {
                var contactId = member.GetAttributeValue<EntityReference>("contact")?.Id ?? Guid.Empty;

                if (contactId != Guid.Empty)
                {
                    Entity contact = service.Retrieve("contact", contactId, new ColumnSet("ntw_isuaeresident"));
                    bool isUaeResident = contact.GetAttributeValue<bool>("ntw_isuaeresident");

                    if (!isUaeResident)
                    {
                        throw new InvalidPluginExecutionException("Manager must be a UAE resident.");
                    }
                }
            }
        }

        private EntityCollection FetchMembers(Entity ifzaLicenseApplication, IOrganizationService service, Guid licenseApplication)
        {
            string fetchXml = $@"
                <fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                    <entity name='ntw_licensecontactmapping'>
                        <attribute name='ntw_contact' />
                        <attribute name='ntw_roles' />
                        <filter type='and'>
                            <condition attribute='ntw_licenseapplication' operator='eq' value='{licenseApplication}' />
                        </filter>
                    </entity>
                </fetch>";

            return service.RetrieveMultiple(new FetchExpression(fetchXml));
        }

        private int CountMembersWithRole(EntityCollection members, string role)
        {
            int count = 0;
            foreach (Entity ent in members.Entities)
            {
                if (ent.Contains("ntw_roles"))
                {
                    Microsoft.Xrm.Sdk.OptionSetValueCollection scc = (Microsoft.Xrm.Sdk.OptionSetValueCollection)ent.Attributes["ntw_roles"];
                    
                    foreach (OptionSetValue opt in scc)
                    {
                        if (opt.Value.ToString() == role)
                            count++;
                    }
                }
            }

           

           

            // Add trace log to print the final count
            tracingService.Trace($"Count of {role}s: {count}");

            return count;
        }

        private bool IsRoleSelected(OptionSetValueCollection optionSetValues, string role)
        {
            // Check if the given role is selected in the multi-option set
            int selectedOptions = GetSelectedOptions(optionSetValues);
            int roleValue = GetRoleValue(role);

            return (selectedOptions & roleValue) == roleValue;
        }

        private int GetSelectedOptions(OptionSetValueCollection optionSetValues)
        {
            int selectedOptions = optionSetValues?.Sum(optionSetValue => optionSetValue.Value) ?? 0;
            return selectedOptions;
        }

        private int GetRoleValue(string role)
        {
            // Define the values corresponding to each role in the multi-option set
            switch (role)
            {
                case "Manager":
                    return 961110000; // Replace with the actual value for "Manager" in the multi-option set
                case "Director":
                    return 961110001; // Replace with the actual value for "Director" in the multi-option set
                case "Shareholder":
                    return 961110002; // Replace with the actual value for "Shareholder" in the multi-option set
                case "Company Secretary":
                    return 961110003; // Replace with the actual value for "Company Secretary" in the multi-option set
                case "Authorized Signatory":
                    return 961110004; // Replace with the actual value for "Authorized Signatory" in the multi-option set
                case "Representative":
                    return 961110005; // Replace with the actual value for "Representative" in the multi-option set
                // Add more roles if needed
                default:
                    return 0;
            }
        }
    }
}
