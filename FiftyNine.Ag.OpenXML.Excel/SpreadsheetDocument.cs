using System;
using FiftyNine.Ag.OpenXML.Common.Packaging;
using FiftyNine.Ag.OpenXML.Common.Parts;
using FiftyNine.Ag.OpenXML.Excel.Parts;

namespace FiftyNine.Ag.OpenXML.Excel
{
    public class SpreadsheetDocument : Package
    {
        public SpreadsheetDocument()
        {
            Workbook = CreatePart<WorkbookPart>("/xl/workbook.xml");
            AppPart = CreatePart<AppPart>("/docProps/app.xml");
            CorePart = CreatePart<CorePart>("/docProps/core.xml");

            AddRelationship(Workbook);
            AddRelationship(AppPart);
            AddRelationship(CorePart);
        }

        public WorkbookPart Workbook
        {
            get;
            private set;
        }
        private AppPart AppPart
        {
            get;
            set;
        }
        private CorePart CorePart
        {
            get;
            set;
        }
        public string ApplicationName
        {
            get { return AppPart.Application; }
            set { AppPart.Application = value; }
        }
        public string Company
        {
            get { return AppPart.Company; }
            set { AppPart.Company = value; }
        }
        public string Creator
        {
            get { return CorePart.Creator; }
            set { CorePart.Creator = value; }
        }
        public DateTime Created
        {
            get { return CorePart.Created; }
            set { CorePart.Created = value; }
        }
    }
}
