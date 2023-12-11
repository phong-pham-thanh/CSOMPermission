using ConsoleCSOM;
using CSOM;
using Microsoft.SharePoint.Client;

string site = "https://sxql8.sharepoint.com/sites/phong-test-sharepoint1/subsite1";
//string site = "https://sxql8.sharepoint.com/sites/phong-test-sharepoint1";
string user = "ptp.phamphong@sxql8.onmicrosoft.com";
string rawPassword = "Phong@123456789";


string termGroupName = "Phong Test group 1";
string termSetName = "Phong Test set 1";
string listTitle = "Account Custom list";
string contentTypeName = "CSOM content type Name";
string fieldAbout = "Field About";
string fieldCity = "Field city";


SharepointManagement sharepointManagement = new(site, user, rawPassword);


try
{
    using (var authenticationManager = new AuthenticationManager())
    using (var context = authenticationManager.GetContext(sharepointManagement.Site, sharepointManagement.Username, sharepointManagement.SecretPassword))
    {
        //await sharepointManagement.CreateListAsync(context, listTitle, Microsoft.SharePoint.Client.ListTemplateType.GenericList);


        //List<string> listAccount = new List<string>();
        //listAccount.Add("phong2");

        //List<string> listPermission = new List<string>();
        //listPermission.Add("Read");
        //listPermission.Add("Contribute");
        //listPermission.Add("Edit");
        //listPermission.Add("Full Control");
        //await sharepointManagement.StopInheritingPermission(context, listTitle);
        //await sharepointManagement.GrantPermissionToList(context, listTitle, listAccount, listPermission);


        //await sharepointManagement.DeleteUniquePermission(context, listTitle);


        //Full Control: Quyền toàn quyền truy cập và quản lý trang SharePoint.
        //Design: Quyền chỉnh sửa cấu trúc trang và nội dung.
        //Edit: Quyền chỉnh sửa nội dung.
        //Contribute: Quyền thêm, sửa, xóa các mục và tài nguyên.
        //Read: Quyền xem nội dung mà không có quyền chỉnh sửa.
        //Limited Access: Quyền truy cập hạn chế, thường được tự động cấp cho người dùng hoặc nhóm khi họ được cấp quyền truy cập tới tài nguyên cụ thể trong trang.
        //View Only: Quyền xem nội dung mà không thể tải về.
        //await sharepointManagement.CreateNewPermissionLevel(context, "Test Permision Level", "This is Description");

        //https://sxql8.sharepoint.com/sites/phong-test-sharepoint1
        //List<string> listAccountMember = new List<string>();
        //listAccountMember.Add("phong2");
        //await sharepointManagement.CreateNewGroup(context, "Test Group Name", "ptp.phamphong", listAccountMember);



        //https://sxql8.sharepoint.com/sites/phong-test-sharepoint1/subsite1
        List<string> listGroupName = new List<string>();
        listGroupName.Add("Test Group Name");
        await sharepointManagement.AssignGroupToList(context, listTitle, listGroupName);












        /////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
        ////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
        ////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
        ////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////













        ////await sharepointManagement.CreateTermGroupAsync(context, termGroupName);
        ////await sharepointManagement.CreateTermSetAsync(context, termGroupName, termSetName);
        ////await sharepointManagement.CreateTermAsync(context, termGroupName, termSetName, "Ho Chi Minh");
        ////await sharepointManagement.CreateTermAsync(context, termGroupName, termSetName, "Stockholm");

        //await sharepointManagement.CreateTextFieldAsync(context, fieldAbout, "About");
        //await sharepointManagement.CreateTaxonomyFieldAsync(context, fieldCity, "city", termGroupName, termSetName);


        //await sharepointManagement.CreateContentTypeAsync(context, contentTypeName, "This is description", null);
        //await sharepointManagement.AddContentTypeToList(context, listTitle, contentTypeName);
        //await sharepointManagement.AddFieldToContentType(context, fieldAbout, contentTypeName);
        //await sharepointManagement.AddFieldToContentType(context, fieldCity, contentTypeName);


        //////await sharepointManagement.SetContentTypeOfListToDefault(context, listTitle, contentTypeName);//Not yet


        //await sharepointManagement.CreateItemToList(context, listTitle, termSetName, new AboutCityStruct { About = "1", City = "Ho Chi Minh" }, fieldAbout, fieldCity);
        //await sharepointManagement.CreateItemToList(context, listTitle, termSetName, new AboutCityStruct { About = "2", City = "Stockholm" }, fieldAbout, fieldCity);
        //await sharepointManagement.CreateItemToList(context, listTitle, termSetName, new AboutCityStruct { About = "3", City = "Ho Chi Minh" }, fieldAbout, fieldCity);
        //await sharepointManagement.CreateItemToList(context, listTitle, termSetName, new AboutCityStruct { About = "4", City = "Stockholm" }, fieldAbout, fieldCity);
        //await sharepointManagement.CreateItemToList(context, listTitle, termSetName, new AboutCityStruct { About = "5", City = "Ho Chi Minh" }, fieldAbout, fieldCity);



        //await sharepointManagement.UpdateFieldAboutToDefaultValue(context, listTitle, fieldAbout, "Default About");
        //await sharepointManagement.CreateItemToList(context, listTitle, termSetName, new AboutCityStruct { About = null, City = "Ho Chi Minh" }, fieldAbout, fieldCity);
        //await sharepointManagement.CreateItemToList(context, listTitle, termSetName, new AboutCityStruct { About = null, City = "Stockholm" }, fieldAbout, fieldCity);


        //await sharepointManagement.UpdateFieldCityToDefaultValue(context, listTitle, termSetName, fieldCity, "Ho Chi Minh");


















    }
}
catch (Exception ex)
{
    Console.WriteLine(1111111 + "    " + ex.Message);
}
