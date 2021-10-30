using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Text;
using ExcelDataReader;
using Newtonsoft.Json;
using System.Text.RegularExpressions;
using System.Data.SqlClient;
using System.Linq;
using System.Threading.Tasks;


namespace FTSExcelProj
{
    class Program
    {
        class MetaDataJsonUser
        {
            public string LastName { get; set; }
            public string FirstName { get; set; }
            public string GeneralInformation { get; set; }
            public string Status { get; set; }
            public string ActivationDate { get; set; }
            public string ExpirationDate { get; set; }
            public string DefaultBadgeForm { get; set; }
            public string ID { get; set; }
            public string HolderInfoClass { get; set; }
            public string LastAdmittedEvent { get; set; }
            public string AccessLevelInformation { get; set; }
            public string Priority { get; set; }
            public string AccessLevel { get; set; }
            public string Activation { get; set; }
            public string Expiration { get; set; }
            public string CustomFieldInformation { get; set; }
            public string CustomFieldName { get; set; }
            public string CustomFieldValue { get; set; }
            public int StartIndex { get; set; }
            public int EndIndex { get; set; }
        }

        class TemplateData
        {
            public IList<ColumnData> ColumnDetails = new List<ColumnData>();
            public IList<FileData> FileDetails = new List<FileData>();
        }

        class FileData
        {
            public string ApplicationName { get; set; }
            public string FileType { get; set; }
            public int StartRowPosition { get; set; }
            public string IgnoreLineswithTags { get; set; }
            public Boolean IsConvertToTextRequired { get; set; }

        }
        class ColumnData
        {
            public string ColumnName { get; set; }
            public string Tag { get; set; }
            public int RowPosition { get; set; }
            public int ColumnValueStartPosition { get; set; }
            public int ColumnValueEndPosition { get; set; }
            public int ColumnValueLength { get; set; }
            public string ColumnValueEndPositionTag { get; set; }
            public int SplitColumnValueIndex { get; set; }
            public string Delimiter { get; set; }
            public string TabularGroupStartTag { get; set; }
            public string TabularGroupEndTag { get; set; }
            public int TabularColumnValueposition { get; set; }
            public bool IsTabularData { get; set; }
            public bool IsFirstColumn { get; set; }
            public bool IsLastColumn { get; set; }
            public bool IsGroupHeader { get; set; }
        }

        class FinalData
        {
            public IList<UserData> UserData = new List<UserData>();
        }

        class UserData
        {
            public string LastName { get; set; }
            public string FirstName { get; set; }
            public string Role { get; set; }
            public string UserName { get; set; }

            public string BU { get; set; }

            public string Asset { get; set; }

            public string AssetDetails { get; set; }

            public string AssetNotes { get; set; }

            public string Status { get; set; }

            public string EmployeeId { get; set; }

            public string TermDate { get; set; }

            public string Email { get; set; }
        }

        static void Main(string[] args)
        {
            string filePathtxt = Path.Combine(Environment.CurrentDirectory);
            int indextxt = filePathtxt.IndexOf("bin");
            TemplateData TemplateObj = new TemplateData();

            string filePathJson = Path.Combine(Environment.CurrentDirectory);
            int indextxtJson = filePathJson.IndexOf("bin");
            //filePathJson = filePathtxt.Substring(0, indextxt - 1) + @"\InputFiles\Irving - xlsx - Template.json";
            filePathJson = filePathtxt.Substring(0, indextxt - 1) + @"\InputFiles\Mainframe - txt - Template.json";
            using (StreamReader jsonfile = File.OpenText(filePathJson))
            {
                JsonSerializer serializer = new JsonSerializer();
                TemplateObj = (TemplateData)serializer.Deserialize(jsonfile, typeof(TemplateData));
            }

            string line;
            List<string> strTextData = new List<string>();
            if (TemplateObj.FileDetails[0].IsConvertToTextRequired)
            {
                System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);
                string filePath = Path.Combine(Environment.CurrentDirectory);
                int index = filePath.IndexOf("bin");
                filePath = filePath.Substring(0, index - 1) + @"\InputFiles\unstrunctal_data.xls";
                filePathtxt = Path.Combine(Environment.CurrentDirectory);
                indextxt = filePathtxt.IndexOf("bin");
                string fileExtension = filePath.Split('.')[1].ToString().ToLowerInvariant();
                filePathtxt = filePathtxt.Substring(0, indextxt - 1) + @"\InputFiles\output.txt";

                using (var stream = File.Open(filePath, FileMode.Open, FileAccess.Read))
                {
                    using (var reader = ExcelReaderFactory.CreateReader(stream))
                    {

                        // 2. Use the AsDataSet extension method
                        DataSet result = reader.AsDataSet();

                        DataTable dt = result.Tables[0];
                        var result2 = CreateDelimitedFileFromDt(dt, ";");
                        FileStream fs;
                        using (fs = File.Create(filePathtxt))
                        {
                            Byte[] temp = new UTF8Encoding(true).GetBytes(result2);
                            fs.Write(temp, 0, temp.Length);
                        }

                    }
                }
                filePathtxt = filePathtxt.Substring(0, indextxt - 1) + @"\InputFiles\output.txt";

                System.IO.StreamReader file = new StreamReader(filePathtxt);
                while ((line = file.ReadLine()) != null)
                {
                    strTextData.Add(line);
                }
            }
            else
            {
                filePathtxt = filePathtxt.Substring(0, indextxt - 1) + @"\InputFiles\DEV8_VRA_NDVR.txt";
                System.IO.StreamReader file = new StreamReader(filePathtxt);
                while ((line = file.ReadLine()) != null)
                {
                    strTextData.Add(line);
                }
            }

            FinalData objFinal = new FinalData();
            UserData userInfo = new UserData();

            bool blGroupStart = false;
            bool blIsTabularColumnHeader = false;
            bool blRecordstart = false;
            bool blIsGroupHeader = false;
            string strGroupHeader = "";

            string strFirstName = "";
            string strLastName = "";
            string strRole = "";

            ColumnData LastNameColumnData = GetSpecificColumnDataFromTemplate(TemplateObj, "LastName");
            ColumnData FirstNameColumnData = GetSpecificColumnDataFromTemplate(TemplateObj, "FirstName");
            ColumnData RoleColumnData = GetSpecificColumnDataFromTemplate(TemplateObj, "Role");
            //ColumnData FullNameColumnData = GetSpecificColumnDataFromTemplate(TemplateObj, "FullName");
            ColumnData UserNameColumnData = GetSpecificColumnDataFromTemplate(TemplateObj, "UserName");

            ColumnData BUColumnData = GetSpecificColumnDataFromTemplate(TemplateObj, "BU");
            ColumnData AssetColumnData = GetSpecificColumnDataFromTemplate(TemplateObj, "Asset");
            ColumnData AssetDetailsColumnData = GetSpecificColumnDataFromTemplate(TemplateObj, "AssetDetails");
            ColumnData AssetNotesColumnData = GetSpecificColumnDataFromTemplate(TemplateObj, "AssetNotes");

            ColumnData TermDateColumnData = GetSpecificColumnDataFromTemplate(TemplateObj, "TermDate");
            ColumnData EmployeeIdColumnData = GetSpecificColumnDataFromTemplate(TemplateObj, "EmployeeId");
            ColumnData EmailColumnData = GetSpecificColumnDataFromTemplate(TemplateObj, "Email");

            //Loop through the lines to parse data
            for (int i = 0; i < strTextData.Count; i++)
            {
                string strCurrentline = strTextData[i];

                // Skip rows as defined in StartRowPosition
                if (i < TemplateObj.FileDetails[0].StartRowPosition)
                    continue;

                // Skip blank lines
                if (strCurrentline.Trim().Length == 0)
                    continue;

                // Skip lines with tags defined in IgnoreLineswithTags
                if (TemplateObj.FileDetails[0].IgnoreLineswithTags.Length > 0)
                {
                    bool blcontinueOuterloop = false;
                    string[] strIgnorestrings = TemplateObj.FileDetails[0].IgnoreLineswithTags.Split(",");
                    for (int k = 0; k < strIgnorestrings.Length; k++)
                    {
                        if (strCurrentline.IndexOf(strIgnorestrings[k]) >= 0)
                        {
                            blcontinueOuterloop = true;
                            continue;
                        }
                    }
                    if (blcontinueOuterloop)
                        continue;
                }

                if (LastNameColumnData != null)
                {
                    if (strCurrentline.IndexOf(LastNameColumnData.Tag) >= 0)
                    {
                        // Row Position is provided, Column value is Above/Below the Tag
                        if (LastNameColumnData.RowPosition != 0 && !LastNameColumnData.IsTabularData)
                        {
                            //SplitColumn
                            if (LastNameColumnData.SplitColumnValueIndex != -1)
                            {
                                int intIndex = i + LastNameColumnData.RowPosition;
                                string strVal = strTextData[intIndex].Substring(LastNameColumnData.ColumnValueStartPosition);
                                string[] strArray = strVal.Split(LastNameColumnData.Delimiter);
                                userInfo.LastName = strArray[LastNameColumnData.SplitColumnValueIndex];
                                strLastName = userInfo.LastName;
                            }
                        }
                        else // Normal Column, Column value beside Left/Right of Tag
                        {

                        }
                    }
                }
                if (FirstNameColumnData != null)
                {
                    if (strCurrentline.IndexOf(FirstNameColumnData.Tag) >= 0)
                    {
                        // Row Position is provided, Column value is Above/Below the Tag row
                        if (FirstNameColumnData.RowPosition != 0 && !FirstNameColumnData.IsTabularData)
                        {
                            //SplitColumn
                            if (FirstNameColumnData.SplitColumnValueIndex != -1)
                            {
                                int intIndex = i + FirstNameColumnData.RowPosition;
                                string strVal = strTextData[intIndex].Substring(FirstNameColumnData.ColumnValueStartPosition);
                                string[] strArray = strVal.Split(FirstNameColumnData.Delimiter);
                                userInfo.FirstName = strArray[FirstNameColumnData.SplitColumnValueIndex];
                                strFirstName = userInfo.FirstName;
                            }
                        }
                        else // Normal Column, Column value beside Left/Right of Tag
                        {

                        }
                    }
                }
                if (RoleColumnData.IsTabularData)
                {
                    if (RoleColumnData.TabularGroupStartTag != null && strCurrentline.IndexOf(RoleColumnData.TabularGroupStartTag) >= 0)
                    {
                        blGroupStart = true;
                        continue;
                    }

                    if (RoleColumnData.IsGroupHeader && strCurrentline.IndexOf(RoleColumnData.Tag) >= 0)
                    {
                        blIsGroupHeader = true;
                        continue;
                    }

                    if (strCurrentline.IndexOf(RoleColumnData.Tag) >= 0 && blGroupStart)
                    {
                        blIsTabularColumnHeader = true;
                        continue;
                    }

                    if (RoleColumnData.TabularGroupEndTag != null && strCurrentline.IndexOf(RoleColumnData.TabularGroupEndTag) >= 0 && blGroupStart)
                    {
                        //blGroupEnd = true;
                        //blRecordend = true;
                        userInfo = new UserData();
                        blIsTabularColumnHeader = false;
                        //blRecordstart = false;
                        //blRecordend = false;
                        blGroupStart = false;
                        //blGroupEnd = false;
                    }

                    if (blIsTabularColumnHeader)
                    {
                        // To pick up Role column value
                        RegexOptions options = RegexOptions.None;
                        Regex regex = new Regex("[;]{2,}", options);
                        strCurrentline = regex.Replace(strCurrentline, ";");
                        if (strCurrentline.IndexOf(';') == 0)
                            strCurrentline = strCurrentline.Substring(1);
                        string[] strAliColumnValues = strCurrentline.Split(';');
                        userInfo.Role = strAliColumnValues[RoleColumnData.TabularColumnValueposition];

                        // Write code to pick up other columns in Tabular data


                        // Add object
                        objFinal.UserData.Add(userInfo);

                        //Initialize user object
                        userInfo = new UserData();

                        //Initialize non repeated column values in new object
                        userInfo.FirstName = strFirstName;
                        userInfo.LastName = strLastName;
                    }

                    // If One role contains multiple users data
                    if (blIsGroupHeader)
                    {

                        if (strCurrentline.Substring(RoleColumnData.ColumnValueStartPosition, RoleColumnData.ColumnValueLength).Trim().Length > 0)
                           strGroupHeader = strCurrentline.Substring(RoleColumnData.ColumnValueStartPosition, RoleColumnData.ColumnValueLength).Trim();

                        strRole = strGroupHeader;
                        userInfo.Role = strGroupHeader;

                        if (UserNameColumnData != null && strCurrentline.IndexOf(UserNameColumnData.Tag) >= 0)
                        {
                            if (UserNameColumnData.ColumnValueLength > 0)
                            {
                                userInfo.UserName = strCurrentline.Substring(UserNameColumnData.ColumnValueStartPosition, UserNameColumnData.ColumnValueLength);
                            }

                            if (UserNameColumnData.ColumnValueEndPositionTag.Length > 0)
                            {
                                int intEndTagIndex = strCurrentline.IndexOf(UserNameColumnData.ColumnValueEndPositionTag) - 1;
                                int intLength = (strCurrentline.IndexOf(UserNameColumnData.ColumnValueEndPositionTag) - 1) - (strCurrentline.IndexOf(UserNameColumnData.Tag) + UserNameColumnData.Tag.Length + UserNameColumnData.ColumnValueStartPosition);
                                userInfo.UserName = strCurrentline.Substring(strCurrentline.IndexOf(UserNameColumnData.Tag)+ UserNameColumnData.Tag.Length+ UserNameColumnData.ColumnValueStartPosition, intLength).Trim();
                            }
                        }

                        if ((FirstNameColumnData != null && strCurrentline.IndexOf(FirstNameColumnData.Tag) >= 0) ||
                            (LastNameColumnData != null && strCurrentline.IndexOf(LastNameColumnData.Tag) >= 0)
                            )
                        {
                            string strName = strCurrentline.Substring(strCurrentline.IndexOf(FirstNameColumnData.Tag)+FirstNameColumnData.Tag.Length+FirstNameColumnData.ColumnValueStartPosition).Trim();
                            string[] strFullname = strName.Split(' ');
                            if (strFullname.Length == 2)
                            {
                                userInfo.LastName = strFullname[1];
                                userInfo.FirstName = strFullname[0];
                            }
                            if (strFullname.Length == 3)
                            {
                                userInfo.LastName = strFullname[2];
                                userInfo.FirstName = strFullname[0];
                            }
                            if (strFullname.Length == 1)
                            {
                                userInfo.FirstName = strFullname[0];
                            }

                        }

                        // Add object
                        objFinal.UserData.Add(userInfo);

                        //Initialize user object
                        userInfo = new UserData();

                        //Initialize non repeated column values in new object
                        userInfo.Role = strRole;
                    }

                }
            } //end for loop

            string jsonN = JsonConvert.SerializeObject(objFinal);

            // Save to Json file
            string OutputfilePath = Path.Combine(Environment.CurrentDirectory);
            int index2 = OutputfilePath.IndexOf("bin");
            OutputfilePath = OutputfilePath.Substring(0, index2 - 1) + @"\OutputFiles\template.json";
            File.WriteAllText(OutputfilePath, jsonN);

            string strJson = jsonN.Substring(jsonN.IndexOf('['));
            strJson = strJson.Substring(0, strJson.Length - 1);
            string ConnectionString = "Server = 10.2.80.81; Database = AMPDEV; user id = ampuser; password = Welcome@123; Integrated Security = false; MultipleActiveResultSets = true; ";
            SqlConnection conn = new SqlConnection(ConnectionString);
            SqlCommand StoredProcedureCommand = new SqlCommand("SaveRFDExData", conn);
            StoredProcedureCommand.CommandType = CommandType.StoredProcedure;
            SqlParameter param1 = new SqlParameter();
            SqlParameter param2 = new SqlParameter();
            SqlParameter param3 = new SqlParameter();
            param1 = StoredProcedureCommand.Parameters.Add("@RFDID", SqlDbType.BigInt);
            param2 = StoredProcedureCommand.Parameters.Add("@DocumentId", SqlDbType.BigInt);
            param3 = StoredProcedureCommand.Parameters.Add("@pJson", SqlDbType.NVarChar, -1);
            param1.Direction = ParameterDirection.Input;
            param2.Direction = ParameterDirection.Input;
            param3.Direction = ParameterDirection.Input;
            param1.Value = 85;
            param2.Value = 72;
            param3.Value = strJson;
            conn.Open();
            SqlDataReader reader1 = StoredProcedureCommand.ExecuteReader();

            while (reader1.Read())
            {
                Console.Write(reader1[0].ToString());
                Console.Write(reader1[1].ToString());
                Console.WriteLine(reader1[2].ToString());
            }
            Console.Write("Successfully Normilized Json Data into Sql Server Table");

            // Close reader and connection
            reader1.Close();
            conn.Close();
        }

        static ColumnData GetSpecificColumnDataFromTemplate(TemplateData td, string strColumn)
        {
            ColumnData columndata = null; //new ColumnData();
            foreach (ColumnData currentColumn in td.ColumnDetails)
            {
                if (currentColumn.ColumnName == strColumn)
                {
                    columndata = currentColumn;
                }
            }
            return columndata;
        }

        /*static void Main(string[] args)
        {
            System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);
            string filePath = Path.Combine(Environment.CurrentDirectory);
            int index = filePath.IndexOf("bin");
            string line;
            filePath = filePath.Substring(0, index - 1) + @"\InputFiles\unstrunctal_data.xls";
            string filePathtxt = Path.Combine(Environment.CurrentDirectory);
            int indextxt = filePathtxt.IndexOf("bin");
            filePathtxt = filePathtxt.Substring(0, indextxt - 1) + @"\InputFiles\output.txt";
            string fileExtension = filePath.Split('.')[1].ToString().ToLowerInvariant();

            using (var stream = File.Open(filePath, FileMode.Open, FileAccess.Read))
            {
                using (var reader = ExcelReaderFactory.CreateReader(stream))
                {

                    // 2. Use the AsDataSet extension method
                    DataSet result = reader.AsDataSet();

                    DataTable dt = result.Tables[0];
                    var result2 = CreateDelimitedFileFromDt(dt, ";");
                    FileStream fs;
                    using (fs = File.Create(filePathtxt))
                    {
                        Byte[] temp = new UTF8Encoding(true).GetBytes(result2);
                        fs.Write(temp, 0, temp.Length);
                    }
                    System.IO.StreamReader file = new StreamReader(filePathtxt);
                    List<string> strTextData = new List<string>();
                    while ((line = file.ReadLine()) != null)
                    {
                        strTextData.Add(line);
                    }
                    string filePathJson = Path.Combine(Environment.CurrentDirectory);
                    int indextxtJson = filePathJson.IndexOf("bin");
                    filePathJson = filePathtxt.Substring(0, indextxt - 1) + @"\InputFiles\UnstructMetadata.json";
                    MetaDataJsonUser MetaDataUser;
                    using (StreamReader jsonfile = File.OpenText(filePathJson))
                    {
                        JsonSerializer serializer = new JsonSerializer();
                        MetaDataUser = (MetaDataJsonUser)serializer.Deserialize(jsonfile, typeof(MetaDataJsonUser));

                    }

                    Data obj = new Data();
                    UserData userData = new UserData();
                    bool isGeneralInfo = false;
                    bool isAccessLevelInfo = false;
                    bool isCustomFieldInfo = false;
                    bool isPriority = false;
                    bool isCustomFieldname = false;

                    //Loop through the lines to parse data
                    for (int i = 0; i < strTextData.Count; i++)
                    {
                        string strCurrentline = strTextData[i];
                        if (strCurrentline.Trim().Length == 0)
                            continue;

                        if (strCurrentline.IndexOf(MetaDataUser.GeneralInformation) > 0)
                        {
                            isGeneralInfo = true;
                            isAccessLevelInfo = false;
                            isCustomFieldInfo = false;

                            string[] strName = strTextData[i - 1].Split(',');
                            userData.LastName = strName[0];
                            userData.FirstName = strName[1];
                            continue;
                        }

                        if (!isGeneralInfo && !isAccessLevelInfo && !isCustomFieldInfo)
                            continue;

                        if (isGeneralInfo)
                        {
                            if (strCurrentline.IndexOf(MetaDataUser.Status) > 0)
                            {
                                int intStatusIndex = strCurrentline.IndexOf(MetaDataUser.Status);
                                int intActivationDateIndex = strCurrentline.IndexOf(MetaDataUser.ActivationDate);
                                int intExpirationDateIndex = strCurrentline.IndexOf(MetaDataUser.ExpirationDate);

                                userData.Status = strCurrentline.Substring(intStatusIndex + MetaDataUser.Status.Length+1, intActivationDateIndex - intStatusIndex - MetaDataUser.Status.Length-1).Trim().Replace(";","");
                                userData.ActivationDate = strCurrentline.Substring(intActivationDateIndex + MetaDataUser.ActivationDate.Length+1, intExpirationDateIndex- intActivationDateIndex - MetaDataUser.ActivationDate.Length-1).Trim().Replace(";", "");
                                userData.ActivationDate = userData.ActivationDate.Substring(MetaDataUser.StartIndex, MetaDataUser.EndIndex);
                                userData.ExpirationDate = strCurrentline.Substring(intExpirationDateIndex + MetaDataUser.ExpirationDate.Length+1).Trim().Replace(";", "");
                                userData.ExpirationDate= userData.ExpirationDate.Substring(MetaDataUser.StartIndex, MetaDataUser.EndIndex);
                            }

                            if (strCurrentline.IndexOf(MetaDataUser.DefaultBadgeForm) > 0)
                            {
                                int intDefaultBadgeFromIndex = strCurrentline.IndexOf(MetaDataUser.DefaultBadgeForm);
                                int intIDIndex = strCurrentline.IndexOf(MetaDataUser.ID);

                                userData.DefaultBadgeForm = strCurrentline.Substring(intDefaultBadgeFromIndex + MetaDataUser.DefaultBadgeForm.Length+1, intIDIndex - intDefaultBadgeFromIndex - MetaDataUser.DefaultBadgeForm.Length-1).Trim().Replace(";", "");
                                userData.ID = strCurrentline.Substring(intIDIndex + MetaDataUser.ID.Length+1).Trim().Replace(";", ""); 
                            }

                            if (strCurrentline.IndexOf(MetaDataUser.HolderInfoClass) > 0)
                            {
                                int intHolderInfoIndex = strCurrentline.IndexOf(MetaDataUser.HolderInfoClass);
                                userData.HolderInfoClass = strCurrentline.Substring(intHolderInfoIndex + MetaDataUser.HolderInfoClass.Length+1).Trim().Replace(";", "");
                            }

                            if (strCurrentline.IndexOf(MetaDataUser.LastAdmittedEvent) > 0)
                            {
                                int intLastAdmittedEventIndex = strCurrentline.IndexOf(MetaDataUser.LastAdmittedEvent);
                                userData.LastAdmittedEvent = strCurrentline.Substring(intLastAdmittedEventIndex + MetaDataUser.LastAdmittedEvent.Length+1).Trim().Replace(";", "");
                            }
                        }

                        if (strCurrentline.IndexOf(MetaDataUser.AccessLevelInformation) > 0)
                        {
                            isAccessLevelInfo = true;
                            isGeneralInfo = false;
                            isCustomFieldInfo = false;
                        }
                        if (isAccessLevelInfo)
                        {
                            if (strCurrentline.IndexOf(MetaDataUser.Priority) > 0)
                            {
                                isPriority = true;
                                continue;
                            }

                            if (isPriority)
                            {
                                RegexOptions options = RegexOptions.None;
                                Regex regex = new Regex("[;]{2,}", options); 
                                strCurrentline = regex.Replace(strCurrentline, ";");
                                if (strCurrentline.IndexOf(';') == 0)
                                    strCurrentline = strCurrentline.Substring(1);
                                string[] strAliColumnValues = strCurrentline.Split(';');

                                AccessLevelInformation ali = new AccessLevelInformation();
                                ali.Priority = strAliColumnValues[0];
                                ali.AccessLevel = strAliColumnValues[1];
                                ali.Activation = strAliColumnValues[2].Substring(MetaDataUser.StartIndex,MetaDataUser.EndIndex);
                                ali.Expiration = strAliColumnValues[3].Substring(MetaDataUser.StartIndex, MetaDataUser.EndIndex); 
                                userData.AccessLevelData.Add(ali);
                            }
                            if (isPriority && (strTextData[i + 1].IndexOf(MetaDataUser.CustomFieldInformation) > 0))
                            {
                                isPriority = false;
                            }
                        }
                        if (strCurrentline.IndexOf(MetaDataUser.CustomFieldInformation) > 0)
                        {
                            isAccessLevelInfo = false;
                            isGeneralInfo = false;
                            isCustomFieldInfo = true;
                        }
                        if (isCustomFieldInfo)
                        {
                            if (strCurrentline.IndexOf(MetaDataUser.CustomFieldName) > 0)
                            {
                                isCustomFieldname = true;
                                continue;
                            }

                            if (isCustomFieldname)
                            {
                                RegexOptions options = RegexOptions.None;
                                Regex regex = new Regex("[;]{2,}", options);
                                strCurrentline = regex.Replace(strCurrentline, ";");
                                if (strCurrentline.IndexOf(';') == 0)
                                    strCurrentline = strCurrentline.Substring(1);
                                string[] strCfiColumnValues = strCurrentline.Split(';');

                                if (strCfiColumnValues.Length == 2)
                                {
                                    CustomFieldInformation cfi = new CustomFieldInformation();
                                    cfi.CustomFieldName = strCfiColumnValues[0];
                                    cfi.CustomFieldValue = strCfiColumnValues[1];
                                    userData.CustomFieldData.Add(cfi);
                                }
                            }
                            if (isCustomFieldname && (
                                CheckData(strTextData, MetaDataUser.GeneralInformation, i + 1)
                                )
                            )
                            {
                                isCustomFieldname = false;
                            }
                        }

                        if ((isCustomFieldInfo &&
                            (CheckData(strTextData, MetaDataUser.GeneralInformation, i + 1)
                            )
                           )
                           ||
                           ((i + 1) > strTextData.Count - 1)
                           )
                        {
                            isAccessLevelInfo = false;
                            isGeneralInfo = false;
                            isCustomFieldInfo = false;
                            obj.UserData.Add(userData);
                            userData = new UserData();
                        }
                    }

                    string jsonN = JsonConvert.SerializeObject(obj);

                    // Save to Json file
                    string OutputfilePath = Path.Combine(Environment.CurrentDirectory);
                    int index2 = OutputfilePath.IndexOf("bin");
                    OutputfilePath = OutputfilePath.Substring(0, index2 - 1) + @"\OutputFiles\template.json";
                    File.WriteAllText(OutputfilePath, jsonN);
        }*/
        //static bool CheckData(List<string> strList, string strTag, int index)
        //{
        //    bool blReturn = false;

        //    if (index > strList.Count - 1)
        //        return false;

        //    if (strList[index].IndexOf(strTag) >= 0)
        //        return true;

        //    return blReturn;
        //}

        static string CreateDelimitedFileFromDt(DataTable dt, string delimiter)
        {
            StringBuilder sb = new StringBuilder();
            foreach (DataRow row in dt.Rows)
            {

                string stuff = "";
                foreach (DataColumn col in row.Table.Columns)
                {
                    string colvalue = Convert.ToString(row[col]);
                    colvalue += delimiter;
                    stuff += colvalue;
                }
                // get rid of delimiter after last column if any
                stuff = stuff.TrimEnd(delimiter.ToCharArray());
                // add line feed
                if (!string.IsNullOrEmpty(stuff))
                {
                    stuff += "\r\n";
                    // append to sb
                    sb.Append(stuff);
                }

            }
            return sb.ToString();
        }
        /*
        class Data
        {
            public IList<UserData> UserData = new List<UserData>();
        }
        class UserData
        {
            public string LastName { get; set; }
            public string FirstName { get; set; }
            public string GeneralInformation { get; set; }
            public string Status { get; set; }
            public string ActivationDate { get; set; }
            public string ExpirationDate { get; set; }
            public string DefaultBadgeForm { get; set; }
            public string ID { get; set; }
            public string HolderInfoClass { get; set; }
            public string LastAdmittedEvent { get; set; }

            public IList<AccessLevelInformation> AccessLevelData = new List<AccessLevelInformation>();
            public IList<CustomFieldInformation> CustomFieldData = new List<CustomFieldInformation>();
        }

        public class AccessLevelInformation
        {
            public string Priority { get; set; }
            public string AccessLevel { get; set; }
            public string Activation { get; set; }
            public string Expiration { get; set; }
        }

        public class CustomFieldInformation
        {
            public string CustomFieldName { get; set; }
            public string CustomFieldValue { get; set; }
        }*/
    }
}
