using Microsoft.BusinessData.MetadataModel;
using Microsoft.SharePoint.Client;
using Microsoft.SharePoint.Client.Taxonomy;
using Microsoft.SharePoint.Marketplace.CorporateCuratedGallery;
using Microsoft.SharePoint.News.DataModel;
using System.Net.Http.Headers;
using System.Security;

namespace CSOM
{
    public struct AboutCityStruct
    {
        public string About;
        public string City;
    }

    public class SharepointManagement
    {




        public Uri Site { get; set; }
        public string Username { get; set; }
        public SecureString SecretPassword { get; set; } = new();


        private readonly int Lcid = 1033; // Locale identifier (LCID) for the language, 1033 is the English language
        private readonly Guid TermStoreId = new Guid("b3680328-f744-4b34-8043-33ed65b18c82");

        public SharepointManagement(string site, string username, string password)
        {
            this.Site = new Uri(site);
            this.Username = username;

            foreach (char c in password.ToCharArray()) SecretPassword.AppendChar(c);
        }
        public async Task UpdateTitleAsync(ClientContext context, string title)
        {
            context.Web.Title = title;
            context.Web.Update();
            await context.ExecuteQueryAsync();
        }

        // sites/precio-homepage/
        public async Task CreateNewWebsiteAsync(ClientContext context, string url, string title)
        {
            WebCreationInformation creation = new();
            creation.Url = url;
            creation.Title = title;
            context.Web.Webs.Add(creation);

            // Retrieve the new web information
            context.Load(context.Web);
            await context.ExecuteQueryAsync();
        }

        public async Task GetAllListsAsync(ClientContext context)
        {
            context.Load(context.Web.Lists);
            await context.ExecuteQueryAsync();

            string text = "";
            foreach (List list in context.Web.Lists)
            {
                text += list.Title + "\n";
            }
            Console.WriteLine(text);
        }

        public async Task GetDocumentsInSubSites(ClientContext context)
        {
            Web rootWeb = context.Site.RootWeb;
            context.Load(rootWeb, w => w.Webs.Include(web => web.ServerRelativeUrl));
            context.ExecuteQuery();

            foreach (var subWeb in rootWeb.Webs)
            {
                Console.WriteLine($"Retrieving documents from sub-site: {subWeb.ServerRelativeUrl}");

                ListCollection lists = subWeb.Lists;
                context.Load(lists, l => l.Include(list => list.BaseType, list => list.Title, list => list.RootFolder.ServerRelativeUrl));
                context.ExecuteQuery();

                foreach (var list in lists)
                {
                    if (list.BaseType == BaseType.DocumentLibrary)
                    {
                        Console.WriteLine($"  Documents in list: {list.Title}");

                        CamlQuery camlQuery = CamlQuery.CreateAllItemsQuery();
                        ListItemCollection items = list.GetItems(camlQuery);
                        context.Load(items, item => item.Include(i => i.File));
                        context.ExecuteQuery();

                        foreach (var item in items)
                        {
                            var file = item.File;
                            Console.WriteLine($"    Document: {file.Name}, URL: {file.ServerRelativeUrl}");
                        }
                    }
                }

                // Recursively retrieve documents from sub-sites
                await GetDocumentsInSubSites(subWeb.Context as ClientContext);
            }
        }

        public async Task CreateListAsync(ClientContext context, string title, ListTemplateType listTemplateType)
        {
            ListCreationInformation creation = new ListCreationInformation();

            creation.Title = title;
            creation.TemplateType = (int)listTemplateType;

            context.Web.Lists.Add(creation);
            await context.ExecuteQueryAsync();
        }

        public async Task UpdateTitleListAsync(ClientContext context, string oldTitle, string newTitle)
        {
            var list = context.Web.Lists.GetByTitle(oldTitle);

            list.Title = newTitle;

            list.Update();
            await context.ExecuteQueryAsync();
        }

        public async Task CreateTermGroupAsync(ClientContext context, string termGroupName)
        {
            TaxonomySession taxonomySession = TaxonomySession.GetTaxonomySession(context);
            context.Load(taxonomySession);
            await context.ExecuteQueryAsync();

            TermStore termStore = taxonomySession.TermStores.GetById(this.TermStoreId);
            context.Load(termStore);
            await context.ExecuteQueryAsync();

            TermGroup termGroup = termStore.CreateGroup(termGroupName, Guid.NewGuid());

            context.Load(termGroup);
            await context.ExecuteQueryAsync();
            Console.WriteLine($"Create term group {termGroupName} successfully.");
        }

        public async Task CreateTermSetAsync(ClientContext context, string termGroupName, string termSetName)
        {
            TaxonomySession taxonomySession = TaxonomySession.GetTaxonomySession(context);
            TermStore termStore = taxonomySession.TermStores.GetById(this.TermStoreId);
            TermGroup termGroup = termStore.Groups.GetByName(termGroupName);
            termGroup.CreateTermSet(termSetName, Guid.NewGuid(), this.Lcid);
            await context.ExecuteQueryAsync();

            Console.WriteLine($"Create term set {termSetName} successfully.");
        }

        public async Task CreateTermAsync(ClientContext context, string termGroupName, string termSetName, string termName)
        {
            TaxonomySession taxonomySession = TaxonomySession.GetTaxonomySession(context);
            TermStore termStore = taxonomySession.TermStores.GetById(this.TermStoreId);
            TermGroup termGroup = termStore.Groups.GetByName(termGroupName);
            TermSet termSet = termGroup.TermSets.GetByName(termSetName);
            termSet.CreateTerm(termName, this.Lcid, Guid.NewGuid());
            await context.ExecuteQueryAsync();

            Console.WriteLine($"Create term {termName} successfully.");
        }

        public async Task<TermSet> GetTermSetAsync(ClientContext context, string termGroupName, string termSetName)
        {
            TaxonomySession taxonomySession = TaxonomySession.GetTaxonomySession(context);
            context.Load(taxonomySession);

            TermStore termStore = taxonomySession.TermStores.GetById(this.TermStoreId);
            context.Load(termStore);

            TermGroup termGroup = termStore.Groups.GetByName(termGroupName);
            context.Load(termGroup);

            TermSet termSet = termGroup.TermSets.GetByName(termSetName);
            context.Load(termSet);

            await context.ExecuteQueryAsync();

            return termSet;
        }


        public async Task CreateTextFieldAsync(ClientContext context, string fieldName, string fieldInternalName)
        {
            FieldCollection siteColumns = context.Web.Fields;
            context.Load(siteColumns);
            Field textField = siteColumns.AddFieldAsXml(
               $"<Field DisplayName='{fieldName}' Type='Text' InternalName='{fieldInternalName}' Group='CSOM City Site Columns' />",
               true,
               AddFieldOptions.DefaultValue);

            await context.ExecuteQueryAsync();
        }


        public async Task CreateTaxonomyFieldAsync(ClientContext context, string fieldName, string fieldInternalName, string termGroupName, string termSetName)
        {
            FieldCollection siteColumns = context.Web.Fields;
            context.Load(siteColumns);

            TermSet termSet = await this.GetTermSetAsync(context, termGroupName, termSetName);

            Field taxonomyField = siteColumns.AddFieldAsXml(
                $"<Field DisplayName='{fieldName}' Type='TaxonomyFieldType' InternalName='{fieldInternalName}' Group='CSOM City Site Columns' />",
                true, AddFieldOptions.DefaultValue);

            // Set additional properties for the taxonomy field
            TaxonomyField taxonomyFieldInstance = context.CastTo<TaxonomyField>(taxonomyField);
            taxonomyFieldInstance.SspId = this.TermStoreId; // Term store ID
            taxonomyFieldInstance.TermSetId = termSet.Id; // Term set ID
            taxonomyFieldInstance.Open = true;
            taxonomyFieldInstance.Update();

            await context.ExecuteQueryAsync();
            Console.WriteLine($"Taxonomy field {fieldName} created successfully.");
        }

        public async Task CreateContentTypeAsync(ClientContext context, string contentTypeName, string contentTypeDescription, string parentContentTypeId)
        {
            ContentTypeCollection contentTypes = context.Web.ContentTypes;
            context.Load(contentTypes);
            ContentTypeCreationInformation contentTypeInfo = new ContentTypeCreationInformation
            {
                Name = contentTypeName,
                Description = contentTypeDescription,
                Group = "CSOM City Content Types",
                ParentContentType = parentContentTypeId != null ? context.Web.ContentTypes.GetById(parentContentTypeId) : null
            };

            ContentType newContentType = contentTypes.Add(contentTypeInfo);
            await context.ExecuteQueryAsync();
        }

        public async Task AddSiteColumnToContentTypeAsync(ClientContext context, string contentTypeName, string siteColumnName)
        {
            context.Load(context.Web);
            await context.ExecuteQueryAsync();

            ContentType contentType = context.Web.ContentTypes.FirstOrDefault(ct => ct.Name == contentTypeName)!;
            context.Load(contentType, ct => ct.FieldLinks, ct => ct.Id);
            await context.ExecuteQueryAsync();

            // Check if the site column is already linked to the content type
            if (!contentType.FieldLinks.Any(fl => fl.Name == siteColumnName))
            {
                // Get the site column
                Field siteColumn = context.Web.Fields.GetByInternalNameOrTitle(siteColumnName);
                context.Load(siteColumn);
                await context.ExecuteQueryAsync();

                // Add the site column to the content type
                FieldLinkCreationInformation fieldLinkInfo = new FieldLinkCreationInformation
                {
                    Field = siteColumn
                };

                contentType.FieldLinks.Add(fieldLinkInfo);
                contentType.Update(true);

                await context.ExecuteQueryAsync();
                Console.WriteLine($"Site column {siteColumnName} added to content type {contentTypeName} successfully.");
            }
            else
            {
                Console.WriteLine($"Site column {siteColumnName} is already linking to content type {contentTypeName}.");
            }
        }


        public async Task AddContentTypeToList(ClientContext context, string listTitle, string contentTypeName)
        {
            List list = context.Web.Lists.GetByTitle(listTitle);
            ContentTypeCollection contentTypes = context.Web.ContentTypes;
            context.Load(list);
            context.Load(contentTypes);
            context.ExecuteQuery();

            ContentType targetContentType = null;
            foreach (var contentType in contentTypes)
            {
                if (contentType.Name == contentTypeName)
                {
                    targetContentType = contentType;
                    break;
                }
            }
            if (targetContentType != null)
            {
                list.ContentTypes.AddExistingContentType(targetContentType);
                context.ExecuteQuery();
                Console.WriteLine("Content type added.");
            }
            else
            {
                Console.WriteLine("Content type not found.");
            }


        }

        public async Task AddFieldToContentType(ClientContext context, string fieldName, string contentTypeName)
        {
            ContentTypeCollection contentTypes = context.Web.ContentTypes;
            FieldCollection fields = context.Web.Fields;

            context.Load(contentTypes);
            context.Load(fields);
            context.ExecuteQuery();

            ContentType targetContentType = null;
            foreach (var contentType in contentTypes)
            {
                if (contentType.Name == contentTypeName)
                {
                    targetContentType = contentType;
                    break;
                }
            }

            Field targetField = null;
            foreach (var field in fields)
            {
                if (field.Title == fieldName)
                {
                    targetField = field;
                    break;
                }
            }


            if (targetContentType != null && targetField != null)
            {
                FieldLinkCreationInformation fieldLink = new FieldLinkCreationInformation();
                fieldLink.Field = targetField;
                targetContentType.FieldLinks.Add(fieldLink);
                targetContentType.Update(true);
                context.ExecuteQuery();
                Console.WriteLine("Field added to content type.");
            }
            else
            {
                Console.WriteLine("Content type or field not found.");
            }
        }

        public async Task SetContentTypeOfListToDefault(ClientContext context, string listTitle, string contentTypeName)
        {
            List list = context.Web.Lists.GetByTitle(listTitle);
            ContentTypeCollection allContentTypes = list.ContentTypes;

            context.Load(allContentTypes);
            context.ExecuteQuery();

            ContentType targetContentType = allContentTypes.FirstOrDefault(ct => ct.Name == contentTypeName);

            if (targetContentType != null)
            {
                // Create a new collection for the content types order
                List<ContentTypeId> contentTypeOrder = new List<ContentTypeId>();

                // Add the target content type id at the first position
                contentTypeOrder.Add(targetContentType.Id);

                // Add the rest of the content types
                foreach (var contentType in list.ContentTypes)
                {
                    if (contentType.Id.StringValue != targetContentType.Id.StringValue)
                    {
                        contentTypeOrder.Add(contentType.Id);
                    }
                }

                // Update the content type order
                list.RootFolder.UniqueContentTypeOrder = contentTypeOrder;
                list.RootFolder.Update();
                list.Update();
                context.ExecuteQuery();
                Console.WriteLine($"'{contentTypeName}' set as default content type for the list '{listTitle}'.");
            }
            else
            {
                Console.WriteLine("Content type not found.");
            }
        }

        public async Task CreateItemToList(ClientContext context, string listTitle, string termSetName, AboutCityStruct item, string FieldNameAbout, string FieldNameCity)
        {
            List list = context.Web.Lists.GetByTitle(listTitle);


            context.Load(list.Fields);
            context.ExecuteQuery();



            TaxonomySession taxonomySession = TaxonomySession.GetTaxonomySession(context);
            TermStore termStore = taxonomySession.GetDefaultSiteCollectionTermStore();
            TermSet termSet = termStore.GetTermSetsByName(termSetName, 1033).GetByName(termSetName);

            // Get the term
            LabelMatchInformation lmi = new LabelMatchInformation(context)
            {
                Lcid = 1033,
                TermLabel = item.City,
                TrimUnavailable = true
            };

            TermCollection termMatches = termSet.GetTerms(lmi);
            context.Load(termMatches);
            context.ExecuteQuery();

            if (termMatches.Count == 0)
            {
                Console.WriteLine("Term not found.");
                return;
            }

            Term term = termMatches.First();


            
            Field fieldContentTypeAbout = list.Fields.Where(fi => fi.Title == FieldNameAbout).FirstOrDefault();
            Field fieldContentTypeCity = list.Fields.Where(fi => fi.Title == FieldNameCity).FirstOrDefault();
            string internalNameAbout = fieldContentTypeAbout != null ? fieldContentTypeAbout.InternalName : "";
            string internalNameCity = fieldContentTypeCity != null ? fieldContentTypeCity.InternalName : "";

            if (internalNameAbout == "" || internalNameCity == "")
            {
                Console.WriteLine("Wwrong Field name");
                return;
            }

            ListItemCreationInformation itemCreateInfo = new ListItemCreationInformation();
            ListItem newItem = list.AddItem(itemCreateInfo);
            newItem["Title"] = "Title";
            if(item.About != null)
            {
                newItem[internalNameAbout] = item.About;
            }
            else
            {
                Field fieldAbout = list.Fields.Where(fi => fi.InternalName == internalNameAbout).FirstOrDefault();
                if (fieldAbout.DefaultValue != null)
                {
                    newItem[internalNameAbout] = fieldAbout.DefaultValue;
                }
            }


            // Set the value for the managed metadata field
            TaxonomyField taxonomyField = context.CastTo<TaxonomyField>(list.Fields.GetByInternalNameOrTitle(internalNameCity));
            taxonomyField.SetFieldValueByTerm(newItem, term, 1033);


            newItem.Update();
            context.ExecuteQuery();
            Console.WriteLine("List items created.");
        }
    
        public async Task UpdateFieldAboutToDefaultValue(ClientContext context, string listTitle, string fieldAbout, string defaultValue)
        {
            // Load the list
            List list = context.Web.Lists.GetByTitle(listTitle);
            context.Load(list);
            context.ExecuteQuery();

            // Retrieve the field
            Field field = list.Fields.GetByInternalNameOrTitle(fieldAbout);
            context.Load(field);
            context.ExecuteQuery();

            // Set the default value
            field.DefaultValue = "Default Value of about";  // The value should be in double quotes
            field.Update();

            // Execute the query
            context.ExecuteQuery();
            Console.WriteLine("Update field About Success");
        }

        public async Task UpdateFieldCityToDefaultValue(ClientContext context, string listTitle, string termSetName, string fieldCity, string defaultValue)
        {


            // Load the list
            List list = context.Web.Lists.GetByTitle(listTitle);
            context.Load(list);
            context.ExecuteQuery();

            // Retrieve the field
            Field field = list.Fields.GetByInternalNameOrTitle(fieldCity);
            context.Load(field);
            context.ExecuteQuery();

            // Set the default value
            field.DefaultValue = "Default Value of about";  // The value should be in double quotes
            field.Update();

            // Execute the query
            context.ExecuteQuery();
            Console.WriteLine("Update field About Success");
























            // Get the list
            //List list = context.Web.Lists.GetByTitle(listTitle);
            //Field field = list.Fields.GetByInternalNameOrTitle(fieldCity);
            //TaxonomyField taxonomyField = context.CastTo<TaxonomyField>(field);
            //context.Load(taxonomyField);
            //context.ExecuteQuery();

            //// Connect to the term store
            //TaxonomySession taxonomySession = TaxonomySession.GetTaxonomySession(context);
            //TermStore termStore = taxonomySession.GetDefaultSiteCollectionTermStore();
            //TermSetCollection termSets = termStore.GetTermSetsByName(termSetName, 1033); // 1033 is the LCID for English
            //context.Load(termSets);
            //context.ExecuteQuery();

            //// Assume single term set for simplicity
            //TermSet termSet = termSets[0];
            //TermCollection terms = termSet.Terms;
            //context.Load(terms);
            //context.ExecuteQuery();

            //// Find the term
            //Term term = terms.FirstOrDefault(t => t.Name == defaultValue);
            //if (term != null)
            //{
            //    // Set the default value
            //    taxonomyField.DefaultValue = term.Id.ToString() + ";#" + term.Name;
            //    taxonomyField.Update();

            //    // Execute the query
            //    context.ExecuteQuery();
            //}
            //else
            //{
            //    Console.WriteLine("Term not found.");
            //}
        }
    
        public async Task GrantPermissionToList(ClientContext context, string listTitle, List<string> accounts, List<string> perrmission)
        {
            List list = context.Web.Lists.GetByTitle(listTitle);
            context.Load(list);
            context.ExecuteQuery();
            foreach(string acc in accounts) {
                string account = acc + "@sxql8.onmicrosoft.com";
                RoleDefinition def;
                RoleDefinitionBindingCollection rdb = new RoleDefinitionBindingCollection(context);
                foreach(string per in perrmission)
                {
                    def = context.Web.RoleDefinitions.GetByName(per);
                    rdb.Add(def);
                }
                Principal usr = context.Web.EnsureUser(account);
                list.RoleAssignments.Add(usr, rdb);
                list.Update();
                context.ExecuteQuery();
                Console.WriteLine("Grant Permission Success");
            }
        }


        public async Task StopInheritingPermission(ClientContext context, string listTitle)
        {
            List list = context.Web.Lists.GetByTitle(listTitle);
            context.Load(list);
            context.ExecuteQuery();

            list.BreakRoleInheritance(true, false);
            list.Update();
            context.ExecuteQuery();
            Console.WriteLine("Stop Inheriting Permission Success");
        }

        public async Task DeleteUniquePermission(ClientContext context, string listTitle)
        {
            List list = context.Web.Lists.GetByTitle(listTitle);
            context.Load(list);
            context.ExecuteQuery();

            list.ResetRoleInheritance();
            list.Update();
            context.ExecuteQuery();
            Console.WriteLine("Delete Unique Permission Success");
        }

        public async Task CreateNewPermissionLevel(ClientContext context, string permissionLevelName, string description)
        {
            Web web = context.Web;
            context.Load(web);
            context.Load(web.AllProperties);
            context.Load(web.RoleDefinitions);
            context.ExecuteQuery();
            var roleDefinitions = web.RoleDefinitions;

            // using role Definition Get by name
            //var fullControlRoleDefinition = roleDefinitions.GetByName("Manage Hierarchy​");
            //context.Load(fullControlRoleDefinition);
            //context.ExecuteQuery();



            // Thiết lập BasePermissions
            BasePermissions permissions = new BasePermissions();
            permissions.Set(PermissionKind.ManageLists);     
            permissions.Set(PermissionKind.CreateAlerts);


            // Create New Custom Permission Level
            RoleDefinitionCreationInformation roleDefinitionCreationInformation = new RoleDefinitionCreationInformation();
            roleDefinitionCreationInformation.BasePermissions = permissions;
            roleDefinitionCreationInformation.Name = permissionLevelName;
            roleDefinitionCreationInformation.Description = description;

            roleDefinitions.Add(roleDefinitionCreationInformation);

            context.Load(roleDefinitions);
            context.ExecuteQuery();
            Console.WriteLine("Create Permission Level Success");
        }
    
        public async Task CreateNewGroup(ClientContext context, string groupName, string ownerName, List<string> membersName)
        {
            Web web = context.Web;
            GroupCreationInformation groupCreationInfo = new GroupCreationInformation();
            groupCreationInfo.Title = groupName;
            groupCreationInfo.Description = "Custom Group Created...";
            User owner = web.EnsureUser(ownerName + "@sxql8.onmicrosoft.com");
            Group group = web.SiteGroups.Add(groupCreationInfo);
            group.Owner = owner;
            foreach(string memberName in membersName)
            {
                User member = web.EnsureUser(memberName + "@sxql8.onmicrosoft.com");
                group.Users.AddUser(member);
            }
            group.Update();
            context.ExecuteQuery();
            Console.WriteLine("Create Group "+ groupName + " Success");

        }

        public async Task AssignGroupToList(ClientContext context, string listTitle, List<string> listGroupName)
        {
            Web web = context.Web;
            List list = web.Lists.GetByTitle(listTitle);
            context.Load(list);
            context.ExecuteQuery();


            // Load the list property(HasUniqueRoleAssignments)
            context.Load(list, target => target.HasUniqueRoleAssignments);
            context.ExecuteQuery();

            if (list.HasUniqueRoleAssignments)
            {
                // Write group name to be added in the list
                Group group = context.Web.SiteGroups.GetByName(listGroupName.FirstOrDefault());
                RoleDefinitionBindingCollection roleDefCollection = new RoleDefinitionBindingCollection(context);

                // Set the permission level of the group for this particular list
                RoleDefinition readDef = context.Web.RoleDefinitions.GetByName("Read");
                roleDefCollection.Add(readDef);

                Principal userGroup = group;
                RoleAssignment roleAssign = list.RoleAssignments.Add(userGroup, roleDefCollection);

                context.Load(roleAssign);
                roleAssign.Update();
                context.ExecuteQuery();
            }
            Console.WriteLine("Assign group success");

        }
    }
}
