using DavigoldExcel.Controls;
using DavigoldExcel.Models;
using DavigoldExcel.Service;
using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Tools.Ribbon;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Net.Http;
using System.Reflection;
using System.Text;
using System.Windows.Forms;

namespace DavigoldExcel
{
    public partial class DavigoldRibbon
    {
        private ImportDataService _importDataService;
        private UploadDataService _uploadDataService;

        private List<dynamic> DownloadFields = new List<dynamic>() {
                new { Module = "Funds",  SubModule = "Fund profile", Required = new List<dynamic>() {
                        new { Slug = "fund-name", Name = "Fund name" }
                    }
                },
                new { Module = "Funds",  SubModule = "Terms", Required = new List<dynamic>() {
                        new { Slug = "fund-name", Name = "Fund name" }
                    }
                },
                new { Module = "Funds",  SubModule = "Strategy", Required = new List<dynamic>() {
                        new { Slug = "fund-name", Name = "Fund name" }
                    }
                },
                new { Module = "Funds",  SubModule = "Shares", Required = new List<dynamic>() {
                        new { Slug = "fund-name", Name = "Fund name" },
                        new { Slug = "share-name", Name = "Share name" },
                    }
                },
                new { Module = "Funds",  SubModule = "Bank Accounts", Required = new List<dynamic>() {
                        new { Slug = "fund-name", Name = "Fund name" }
                    }
                },
                new { Module = "Funds",  SubModule = "Closings", Required = new List<dynamic>() {
                        new { Slug = "fund-name", Name = "Fund name" }
                    }
                },
                new { Module = "Funds",  SubModule = "NAV", Required = new List<dynamic>() {
                        new { Slug = "fund-name", Name = "Fund name" },
                        new { Slug = "date", Name = "Date" },
                    }
                },
                new { Module = "Funds",  SubModule = "NAV Breakdown", Required = new List<dynamic>() {
                        new { Slug = "fund-name", Name = "Fund name" },
                        new { Slug = "share-name", Name = "Share name" },
                        new { Slug = "lp-last-name", Name = "LP Last name" },
                        new { Slug = "lp-first-name", Name = "LP First name" },
                    }
                },
                new { Module = "Funds",  SubModule = "Accounts", Required = new List<dynamic>() {
                        new { Slug = "fund-name", Name = "Fund name" },
                    }
                },
                new { Module = "Funds",  SubModule = "Accounts groups", Required = new List<dynamic>() {
                        new { Slug = "fund-name", Name = "Fund name" },
                    }
                },
                new { Module = "Limited Partners",  SubModule = "LP Details", Required = new List<dynamic>() {
                        new { Slug = "lp-last-name", Name = "LP Last name" },
                        new { Slug = "lp-first-name", Name = "LP First name" },
                    }
                },
                new { Module = "Limited Partners",  SubModule = "Commitments", Required = new List<dynamic>() {
                        new { Slug = "to-lp-last-name", Name = "LP Last name" },
                        new { Slug = "to-lp-first-name", Name = "LP First name" },
                        new { Slug = "fund-name", Name = "Fund name" },
                        new { Slug = "share-name", Name = "Share" },
                        new { Slug = "date", Name = "Date" },
                    }
                },
                new { Module = "Limited Partners",  SubModule = "LP Operations", Required = new List<dynamic>() {
                        //new { Slug = "lp-last-name", Name = "LP Last name" },
                        //new { Slug = "lp-first-name", Name = "LP First name" },
                        //new { Slug = "fund-name", Name = "Fund name" },
                        //new { Slug = "share", Name = "Share" },
                        new { Slug = "date", Name = "Date" },
                    }
                },
                new { Module = "Limited Partners",  SubModule = "Capital Calls", Required = new List<dynamic>() {
                        new { Slug = "fund", Name = "Fund name" },
                        new { Slug = "call-name", Name = "Call name" },
                    }
                },
                new { Module = "Limited Partners",  SubModule = "Distributions", Required = new List<dynamic>() {
                        new { Slug = "fund", Name = "Fund name" },
                        new { Slug = "distribution-name", Name = "Distribution name" },
                    }
                },
                new { Module = "Portfolio",  SubModule = "Portfolio high level", Required = new List<dynamic>() {
                        new { Slug = "portfolio-name", Name = "Deal name" },
                        new { Slug = "portfolio-company", Name = "Portfolio company" },
                    }
                },
                new { Module = "Portfolio",  SubModule = "Details", Required = new List<dynamic>() {
                        new { Slug = "portfolio-name", Name = "Deal name" },
                        new { Slug = "portfolio-company", Name = "Portfolio company" },
                    }
                },
                new { Module = "Portfolio",  SubModule = "Fund info", Required = new List<dynamic>() {
                        new { Slug = "portfolio-name", Name = "Deal name" },
                        new { Slug = "linked-to-funds", Name = "Fund name" },
                    }
                },
                new { Module = "Portfolio",  SubModule = "Legal", Required = new List<dynamic>() {
                        new { Slug = "portfolio-company", Name = "Portfolio company" },
                    }
                },
                new { Module = "Portfolio",  SubModule = "Key figures", Required = new List<dynamic>() {
                        new { Slug = "portfolio-name", Name = "Portfolio company" },
                        new { Slug = "table", Name = "Table" },
                        new { Slug = "kpi", Name = "KPI" },
                    }
                },
                new { Module = "Portfolio",  SubModule = "Shareholders", Required = new List<dynamic>() {
                        new { Slug = "portfolio-company", Name = "Portfolio company" },
                    }
                },
                new { Module = "Portfolio",  SubModule = "Operations", Required = new List<dynamic>() {
                        new { Slug = "portfolio-company", Name = "Portfolio company" },
                        new { Slug = "fund-name", Name = "Fund name" },
                        new { Slug = "date", Name = "Date" },
                    }
                },
                new { Module = "Portfolio",  SubModule = "Quarterly Updates", Required = new List<dynamic>() {
                        new { Slug = "portfolio-company", Name = "Portfolio company" },
                        new { Slug = "header", Name = "Header" },
                        new { Slug = "description", Name = "Description" },
                    }
                },
                new { Module = "Portfolio",  SubModule = "Securities", Required = new List<dynamic>() {
                        new { Slug = "portfolio-company", Name = "Portfolio company" },
                        new { Slug = "security-type", Name = "Security type" },
                        new { Slug = "name", Name = "Name" },
                    }
                },
                new { Module = "Portfolio",  SubModule = "Valuations", Required = new List<dynamic>() {
                        new { Slug = "portfolio-company", Name = "Portfolio company" },
                        new { Slug = "fund-name", Name = "Fund name" },
                        new { Slug = "security-name", Name = "Security name" },
                    }
                },
                new { Module = "Deal flow",  SubModule = "Deals", Required = new List<dynamic>() {
                        new { Slug = "deal-name", Name = "Deal name" },
                        new { Slug = "main-target", Name = "Main target" },
                    }
                },
                new { Module = "Deal flow",  SubModule = "Details", Required = new List<dynamic>() {
                        new { Slug = "deal-name", Name = "Deal name" },
                        new { Slug = "main-target", Name = "Main target" },
                    }
                },
                new { Module = "Deal flow",  SubModule = "Fund info", Required = new List<dynamic>() {
                        new { Slug = "deal-name", Name = "Deal name" },
                        new { Slug = "linked-to-funds", Name = "Fund name" },
                    }
                },
                new { Module = "Deal flow",  SubModule = "Stages", Required = new List<dynamic>() {
                        new { Slug = "deal-name", Name = "Deal name" },
                    }
                },
                new { Module = "Deal flow",  SubModule = "Deal Contacts", Required = new List<dynamic>() {
                        new { Slug = "deal-name", Name = "Deal name" },
                    }
                },
                new { Module = "Companies",  SubModule = "Companies", Required = new List<dynamic>() {
                        new { Slug = "company-name", Name = "Company name" },
                    }
                },
                new { Module = "Companies",  SubModule = "Offices", Required = new List<dynamic>() {
                        new { Slug = "company-name", Name = "Company name" },
                        new { Slug = "country", Name = "Country" },
                    }
                },
                new { Module = "Companies",  SubModule = "Executive Team", Required = new List<dynamic>() {
                        new { Slug = "company-name", Name = "Company name" },
                    }
                },
                new { Module = "Contacts",  SubModule = "Contacts", Required = new List<dynamic>() {
                        new { Slug = "last-name", Name = "Last name" },
                        new { Slug = "first-name", Name = "First name" },
                    }
                },
                new { Module = "Contacts",  SubModule = "Positions", Required = new List<dynamic>() {
                        new { Slug = "last-name", Name = "Last name" },
                        new { Slug = "first-name", Name = "First name" },
                        new { Slug = "position-company-name", Name = "Company name" },
                    }
                },
                new { Module = "Administration",  SubModule = "Currencies", Required = new List<dynamic>() {
                        new { Slug = "admin-currencies-date", Name = "Date" },
                        new { Slug = "admin-currencies-reference-currency", Name = "Reference currency" },
                        new { Slug = "admin-currencies-fx-currency", Name = "FX currency" },
                        new { Slug = "admin-currencies-rate", Name = "Rate" },
                    }
                },
        };

        private List<dynamic> UploadFields = new List<dynamic>() {
                new { Module = "Funds",  SubModule = "Fund profile", Required = new List<dynamic>() {
                        new { Slug = "fund-name", Name = "Fund name" },
                        new { Slug = "general-partner", Name = "General Partner" },
                        new { Slug = "currency", Name = "Currency" },
                        new { Slug = "formation-on", Name = "Formation on" },
                    }
                },
                new { Module = "Funds",  SubModule = "Terms", Required = new List<dynamic>() {
                        new { Slug = "fund-name", Name = "Fund name" },
                        new { Slug = "country-of-registration", Name = "Country of registration" },
                        new { Slug = "legal-form", Name = "Legal form" },
                    }
                },
                new { Module = "Funds",  SubModule = "Strategy", Required = new List<dynamic>() {
                        new { Slug = "fund-name", Name = "Fund name" }
                    }
                },
                new { Module = "Funds",  SubModule = "Shares", Required = new List<dynamic>() {
                        new { Slug = "fund-name", Name = "Fund name" },
                        new { Slug = "share-name", Name = "Share name" },
                        new { Slug = "nominal-value", Name = "Nominal value" },
                    }
                },
                new { Module = "Funds",  SubModule = "Bank Accounts", Required = new List<dynamic>() {
                        new { Slug = "fund-name", Name = "Fund name" },
                        new { Slug = "iban", Name = "IBAN" }
                    }
                },
                new { Module = "Funds",  SubModule = "Closings", Required = new List<dynamic>() {
                        new { Slug = "fund-name", Name = "Fund name" },
                        new { Slug = "closing-number", Name = "Closing number" },
                        new { Slug = "closing-start-date", Name = "Start date" },
                        new { Slug = "closing-end-date", Name = "End date" },
                    }
                },
                new { Module = "Funds",  SubModule = "NAV", Required = new List<dynamic>() {
                        new { Slug = "fund-name", Name = "Fund name" },
                        new { Slug = "date", Name = "Date" },
                        new { Slug = "total-fund-nav", Name = "Fund NAV" },
                    }
                },
                new { Module = "Funds",  SubModule = "NAV Breakdown", Required = new List<dynamic>() {
                        new { Slug = "fund-name", Name = "Fund name" },
                        new { Slug = "share-name", Name = "Share name" },
                        new { Slug = "date", Name = "Date" },
                        new { Slug = "lp-id", Name = "LP ID" },
                        new { Slug = "amount-breakdown", Name = "Amount" },
                    }
                },
                new { Module = "Funds",  SubModule = "Accounts", Required = new List<dynamic>() {
                        new { Slug = "fund-name", Name = "Fund name" },
                    }
                },
                new { Module = "Funds",  SubModule = "Accounts groups", Required = new List<dynamic>() {
                        new { Slug = "fund-name", Name = "Fund name" },
                    }
                },
                new { Module = "Limited Partners",  SubModule = "LP Details", Required = new List<dynamic>() {
                        new { Slug = "lp-id", Name = "LP ID" },
                        new { Slug = "lp-last-name", Name = "LP Last name" },
                    },
                    ConditionalRequired = new List<dynamic>() {
                    new {
                        If = "share-name",
                        Then = new List<dynamic>() {
                            new { Slug = "fund-name", Name = "Fund name" },
                            new { Slug = "lp-details-date", Name = "Date" },
                            new { Slug = "lp-details-no-of-shares", Name = "No. of shares" }
                        }
                    }
                }
                },
                new { Module = "Limited Partners",  SubModule = "Commitments", Required = new List<dynamic>() {
                    new { Slug = "lp-id", Name = "LP ID" },
                    new { Slug = "commitment-id", Name = "Commitment ID" },

                    }
                },
                new { Module = "Limited Partners",  SubModule = "Commitment Info", Required = new List<dynamic>() {
                        new { Slug = "lp-last-name", Name = "LP Last name" },
                        new { Slug = "lp-first-name", Name = "LP First name" },
                        new { Slug = "fund-name", Name = "Fund name" },
                        new { Slug = "commitment-date", Name = "Commitment Date" },
                    }
                },
                new { Module = "Limited Partners",  SubModule = "LP Operations", Required = new List<dynamic>() {
                        new { Slug = "lp-last-name", Name = "LP Last name" },
                        new { Slug = "lp-first-name", Name = "LP First name" },
                        new { Slug = "fund-name", Name = "Fund name" },
                        new { Slug = "Share", Name = "Share" },
                        new { Slug = "date", Name = "Date" },
                        new { Slug = "operation-type", Name = "Operation type" },
                    }
                },
                new { Module = "Limited Partners",  SubModule = "Capital Calls", Required = new List<dynamic>() {
                        new { Slug = "fund", Name = "Fund name" },
                        new { Slug = "call-name", Name = "Call name" },
                        new { Slug = "call-date", Name = "Call date" },
                        new { Slug = "due-date", Name = "Due date" },
                        new { Slug = "share-name", Name = "Share name" },
                        new { Slug = "closing", Name = "Closing" },
                        new { Slug = "percentage-called", Name = "% Called" },
                    }
                },
                new { Module = "Limited Partners",  SubModule = "Capital Calls Breakdown", Required = new List<dynamic>() {
                        new { Slug = "fund-name", Name = "Fund name" },
                        new { Slug = "call-name", Name = "Call name" },
                        new { Slug = "date", Name = "Due date" },
                        new { Slug = "lp-id", Name = "LP ID" }, 
                        new { Slug = "share-name", Name = "Share name" },
                    }
                },
                new { Module = "Limited Partners",  SubModule = "Distributions", Required = new List<dynamic>() {
                        new { Slug = "fund", Name = "Fund name" },
                        new { Slug = "distribution-name", Name = "Distribution name" },
                        new { Slug = "date", Name = "Distribution date" },
                        new { Slug = "share-name", Name = "Share name" },
                        new { Slug = "amount", Name = "Distributed amount" },
                    }
                },
                new { Module = "Limited Partners",  SubModule = "Distributions Breakdown", Required = new List<dynamic>() {
                        new { Slug = "fund-name", Name = "Fund name" },
                        new { Slug = "distribution-name", Name = "Distribution name" },
                        new { Slug = "date", Name = "Due date" },
                        new { Slug = "lp-id", Name = "LP ID" },
                        new { Slug = "share-name", Name = "Share name" },
                        new { Slug = "amount-breakdown", Name = "Amount Breakdown" },
                    }
                },
                new { Module = "Portfolio",  SubModule = "Portfolio high level", Required = new List<dynamic>() {
                        new { Slug = "deal-id", Name = "Deal ID" },
                        new { Slug = "portfolio-name", Name = "Deal name" },
                        new { Slug = "portfolio-company", Name = "Portfolio company" },
                        new { Slug = "currency", Name = "Currency" },
                        new { Slug = "linked-to-funds", Name = "Linked to funds" },
                    }
                },
                new { Module = "Portfolio",  SubModule = "Details", Required = new List<dynamic>() {
                        new { Slug = "portfolio-name", Name = "Deal name" },
                        new { Slug = "portfolio-company", Name = "Portfolio company" },
                    }
                },
                new { Module = "Portfolio",  SubModule = "Fund info", Required = new List<dynamic>() {
                        new { Slug = "deal-id", Name = "Deal ID" },
                        new { Slug = "linked-to-funds", Name = "Fund name" }
                    }
                },
                new { Module = "Portfolio",  SubModule = "Portfolio Contacts", Required = new List<dynamic>() {
                        new { Slug = "deal-id", Name = "Deal ID" },
                        new { Slug = "portfolio-team", Name = "Team Member" },
                    }
                },
                new { Module = "Portfolio",  SubModule = "Legal", Required = new List<dynamic>() {
                        new { Slug = "portfolio-company", Name = "Portfolio company" }
                    }
                },
                new { Module = "Portfolio",  SubModule = "Key figures", Required = new List<dynamic>() {
                        new { Slug = "portfolio-name", Name = "Portfolio company" },
                        new { Slug = "linked-to-funds", Name = "Linked funds" },
                        new { Slug = "key-figure-as-of", Name = "Key figure as of" },
                        new { Slug = "table", Name = "Table" },
                        new { Slug = "kpi", Name = "KPI" },
                        new { Slug = "value", Name = "Value" },
                        new { Slug = "year", Name = "Year" },
                    }
                },
                new { Module = "Portfolio",  SubModule = "Shareholders", Required = new List<dynamic>() {
                        new { Slug = "deal-id", Name = "Deal Id" },
                        new { Slug = "portfolio-company", Name = "Portfolio company" },
                        new { Slug = "linked-to-funds", Name = "Linked funds" },
                        new { Slug = "capital-table-date", Name = "Capital table date" },
                        new { Slug = "shareholder-type", Name = "Shareholder type" },
                        new { Slug = "shareholder-last-name", Name = "Last name" },
                        new { Slug = "shareholder-first-name", Name = "First name" },
                        new { Slug = "non-diluted", Name = "% Non diluted" },
                    }
                },
                new { Module = "Portfolio",  SubModule = "Operations", Required = new List<dynamic>() {
                        new { Slug = "portfolio-company", Name = "Portfolio company" },
                        new { Slug = "deal-name", Name = "Deal name" },
                        new { Slug = "date", Name = "Date" },
                        new { Slug = "type", Name = "Type" },
                        new { Slug = "amount", Name = "Amount" },
                        new { Slug = "amount-fx", Name = "Amount FX" },
                        new { Slug = "from-last-name", Name = "From Last name" },
                        new { Slug = "from-first-name", Name = "From First name" },
                        new { Slug = "to-last-name", Name = "To Last name" },
                        new { Slug = "to-first-name", Name = "To First name" },
                    }
                },
                new { Module = "Portfolio",  SubModule = "Quarterly Updates", Required = new List<dynamic>() {
                        new { Slug = "deal-id", Name = "Deal Id" },
                        new { Slug = "date", Name = "Date" },
                        new { Slug = "header", Name = "Header" },
                        new { Slug = "description", Name = "Description" },
                    }
                },
                new { Module = "Portfolio",  SubModule = "Securities", Required = new List<dynamic>() {
                        new { Slug = "deal-id" , Name = "Deal Id" },
                        new { Slug = "security-type", Name = "Security type" },
                        new { Slug = "name", Name = "Name" },
                        new { Slug = "nominal-value", Name = "Nominal value" },
                    }
                },
                new { Module = "Portfolio",  SubModule = "Valuations", Required = new List<dynamic>() {
                        new { Slug = "deal-id" , Name = "Deal Id" },
                        new { Slug = "fund-name", Name = "Fund name" },
                        new { Slug = "date", Name = "Date" },
                    }
                },
                new { Module = "Portfolio",  SubModule = "Assets", Required = new List<dynamic>() {
                        new { Slug = "company-asset-main-company-name" , Name = "Portfolio Company name" },
                        new { Slug = "company-asset-name", Name = "Asset name" },
                    }
                },
                new { Module = "Deal flow",  SubModule = "Deals", Required = new List<dynamic>() {
                        new { Slug = "deal-id", Name = "Deal ID" },
                        new { Slug = "currency", Name = "Currency" },
                        new { Slug = "deal-last-stage", Name = "Deal last stage" },
                    }
                },
                new { Module = "Deal flow",  SubModule = "Details", Required = new List<dynamic>() {
                        new { Slug = "deal-id", Name = "Deal ID" },
                        new { Slug = "deal-name", Name = "Deal name" },
                        new { Slug = "main-target", Name = "Main target" },
                        new { Slug = "linked-to-funds", Name = "Linked funds" },
                    }
                },
                new { Module = "Deal flow",  SubModule = "Fund info", Required = new List<dynamic>() {
                        new { Slug = "deal-id", Name = "Deal ID" },
                        new { Slug = "linked-to-funds", Name = "Fund name" }
                    }
                },
                new { Module = "Deal flow",  SubModule = "Stages", Required = new List<dynamic>() {
                        new { Slug = "deal-id", Name = "Deal ID" },
                        new { Slug = "deal-name", Name = "Deal name" },
                        new { Slug = "date", Name = "Date" },
                        new { Slug = "Stage", Name = "Stage" },
                        new { Slug = "linked-to-funds", Name = "Linked funds" },
                    }
                },
                new { Module = "Deal flow",  SubModule = "Deal Contacts", Required = new List<dynamic>() {
                        new { Slug = "deal-id", Name = "Deal ID" },
                        new { Slug = "deal-team", Name = "Team Member" },
                    }
                },
                new { Module = "Companies",  SubModule = "Companies", Required = new List<dynamic>() {
                        new { Slug = "custom-id", Name = "ID" },
                        new { Slug = "company-name", Name = "Company name" },
                        new { Slug = "country", Name = "Country" },

                    }
                },
                new { Module = "Companies",  SubModule = "Offices", Required = new List<dynamic>() {
                        new { Slug = "company-name", Name = "Company name" },
                        new { Slug = "country", Name = "Country" },
                    }
                },
                new { Module = "Companies",  SubModule = "Executive Team", Required = new List<dynamic>() {
                        new { Slug = "company-name", Name = "Company name" },
                        new { Slug = "contact-last-name", Name = "Contact Last name" },
                        new { Slug = "contact-first-name", Name = "Contact First name" },
                        new { Slug = "team-type", Name = "Team type" },
                    }
                },
                new { Module = "Contacts",  SubModule = "Contacts", Required = new List<dynamic>() {
                        new { Slug = "last-name", Name = "Last name" },
                        new { Slug = "first-name", Name = "First name" },
                        new { Slug = "custom-id", Name = "ID" },
                    }
                },
                new { Module = "Contacts",  SubModule = "Positions", Required = new List<dynamic>() {
                        new { Slug = "last-name", Name = "Last name" },
                        new { Slug = "first-name", Name = "First name" },
                        new { Slug = "position-company-name", Name = "Company name" },
                        new { Slug = "position-team-type", Name = "Team type" },
                    }
                },
               new { Module = "Administration",  SubModule = "Currencies", Required = new List<dynamic>() {
                        new { Slug = "admin-currencies-date", Name = "Date" },
                        new { Slug = "admin-currencies-reference-currency", Name = "Reference currency" },
                        new { Slug = "admin-currencies-fx-currency", Name = "FX currency" },
                        new { Slug = "admin-currencies-rate", Name = "Rate" },
                    }
                },
        };

        private void DavigoldRibbon_Load(object sender, RibbonUIEventArgs e)
        {
            _importDataService = new ImportDataService();
            _uploadDataService = new UploadDataService();

            var version = typeof(ThisAddIn).Assembly.GetName().Version;
            DateTime _dateTime = new DateTime(2025, 10, 13);
            VersionLabel.Label = $"Version: {version}";
            UpdatedOnLabel.Label = $"Last Updated on {_dateTime.ToString("dd MMMM yyyy", CultureInfo.InvariantCulture)}";
            //var currentId = ExcelService.GetStoreDataInWoorkbook("tenantId");
            if (!UTILS.showSidePanel)
            {
                ShowHideButton.Visible = false;
            }
        }

        private async void ImportButton_Click(object sender, RibbonControlEventArgs e)
        {
            Microsoft.Office.Interop.Excel.Application excelApp = Globals.ThisAddIn.Application;
            Worksheet activeSheet = excelApp.ActiveSheet as Worksheet;

            Range usedRange = activeSheet.UsedRange;

            int lastUsedRow = 1;
            int lastUsedCol = 1;

            Microsoft.Office.Interop.Excel.Comments comments = activeSheet.Comments;
            if (comments != null && comments.Count > 0)
            {
                for (int i = 1; i <= comments.Count; i++)
                {
                    Comment comment = comments[i];
                    if (comment?.Parent is Range cell)
                    {
                        int row = cell.Row;
                        int col = cell.Column;

                        if (row > lastUsedRow) lastUsedRow = row;
                        if (col > lastUsedCol) lastUsedCol = col;
                    }
                }
            }

            //int lastUsedRow = activeSheet.Cells.Find("*", Missing.Value, Missing.Value, Missing.Value, XlSearchOrder.xlByRows, XlSearchDirection.xlPrevious, false, Missing.Value, Missing.Value)?.Row ?? 1;
            //int lastUsedCol = activeSheet.Cells.Find("*", Missing.Value, Missing.Value, Missing.Value, XlSearchOrder.xlByColumns, XlSearchDirection.xlPrevious, false, Missing.Value, Missing.Value)?.Column ?? 1;

            Cursor.Current = Cursors.WaitCursor;
            excelApp.Interactive = false;


            for (int i = 1; i <= lastUsedRow; i++)
            {
                bool hasTable = false;
                bool hasValuation = false;
                string module = null;
                string subModule = null;
                List<dynamic> columns = new List<dynamic>();

                for (int j = 1; j <= lastUsedCol; j++)
                {
                    Range currentRange = activeSheet.Cells[i, j] as Range;

                    string cellComment = currentRange.Comment != null ? currentRange.Comment.Text() : null;
                    string cellValue = currentRange.Value2 as string;
                    if (cellComment != null && !String.IsNullOrEmpty(cellComment))
                    {
                        List<string> commentData = cellComment.Split(':').ToList();
                        if (commentData.Count >= 5)
                        {
                            if (commentData[0] == "D" && commentData[4] == "list")
                            {
                                if (!hasTable)
                                {
                                    hasTable = true;
                                }

                                if (module == null)
                                {
                                    module = commentData[1];
                                }

                                if (subModule == null)
                                {
                                    subModule = commentData[2];
                                }

                                if (hasValuation == false && module == "Portfolio" && commentData[2] == "Valuation")
                                {
                                    hasValuation = true;
                                }

                                columns.Add(new { Name = commentData[3], Position = j });
                            }

                        }
                        else if (hasTable && cellComment.Contains("-kpi:list"))
                        {
                            columns.Add(new { Name = cellComment, Position = j });
                        }
                    }
                }

                if (hasTable && columns.Count > 0)
                {
                    List<string> currentColumns = columns.Select(column => (string)column.Name).ToList();
                    if (hasValuation)
                    {
                        currentColumns.Add("has-valuation");
                    }
                    string errorMessage = ValidateDownload(module, subModule, currentColumns);
                    if (errorMessage == null)
                    {
                        MessageBox.Show(
                             "The download has started",
                             "Downloading",
                             MessageBoxButtons.OK,
                             MessageBoxIcon.Information
                        );
                        string currentFund = Home.Instance.SelectedFundComboBoxItem.Value;
                        string language = Home.Instance.IsEnglish ? "EN" : "FR";
                        string filterDate = null;
                        
                        // Get the date from FilterDatePicker in the Home control
                        var taskPane = Globals.ThisAddIn.GetCurrentTaskPane();
                        if (taskPane != null && taskPane.Control is MainWindow mainWindow)
                        {
                            var homeControl = mainWindow.mainFormHost.Child as Home;
                            if (homeControl != null)
                            {
                                var datePicker = homeControl.GetFilterDatePicker();
                                if (datePicker != null && datePicker.SelectedDate.HasValue)
                                {
                                    filterDate = datePicker.SelectedDate.Value.ToString("yyyy-MM-dd");
                                }
                            }
                        }

                        try
                        {
                            List<dynamic> data = await _importDataService.ImportData<dynamic>(currentColumns, currentFund, module, subModule, language, filterDate);
                            if (data != null && data.Count > 0)
                            {
                                ExcelService.InsertImportedData(data, columns, i + 1);

                                lastUsedRow += data.Count;
                                i += data.Count;
                            }
                        }
                        catch (HttpRequestException)
                        {
                        }
                    }
                    else
                    {
                        MessageBox.Show(
                           $"Please Insert all the mandatory fields before downloading ({errorMessage})",
                           "Required Fields",
                           MessageBoxButtons.OK,
                           MessageBoxIcon.Error
                       );
                    }
                }
            }

            excelApp.Interactive = true;
            Cursor.Current = Cursors.Default;

            activeSheet.Cells.EntireColumn.AutoFit();
            activeSheet.Cells.EntireRow.AutoFit();



            //string CurrentSheetModule = ExcelService.GetStoreDataInActiveSheet("Module");
            //string CurrentSheetSubModule = ExcelService.GetStoreDataInActiveSheet("SubModule");

            //List<string> Columns = ExcelService.GetColumns();

            //if (CurrentSheetModule != null && CurrentSheetSubModule != null && Columns.Count > 0)
            //{
            //    Microsoft.Office.Interop.Excel.Application excelApp = Globals.ThisAddIn.Application;

            //    excelApp.Interactive = false;

            //    Cursor.Current = Cursors.WaitCursor;
            //    List<dynamic> data = await _importDataService.ImportData<dynamic>(Columns, CurrentSheetModule, CurrentSheetSubModule);
            //    if (data != null)
            //    {
            //        ExcelService.InsertImportedData(data);

            //        Worksheet worksheet = excelApp.ActiveSheet as Worksheet;

            //        worksheet.Cells.EntireColumn.AutoFit();
            //        worksheet.Cells.EntireRow.AutoFit();

            //        Cursor.Current = Cursors.Default;
            //        //MessageBox.Show("Data successfully imported into the sheet.", "Data Imported", System.Windows.MessageBoxButton.OK, System.Windows.MessageBoxImage.Information);
            //    }
            //    else
            //    {
            //        Cursor.Current = Cursors.Default;
            //    }

            //    excelApp.Interactive = true;
            //}
        }
        private bool HasProperty(object obj, string propertyName)
        {
            return TypeDescriptor.GetProperties(obj).Find(propertyName, ignoreCase: true) != null;
        }

        private string ValidateDownload(string module, string subModule, List<string> fields)
        {
            dynamic currentModule = DownloadFields.FirstOrDefault(d => d.Module == module && d.SubModule == subModule);
            if (currentModule != null)
            {
                List<dynamic> RequiredFields = currentModule.Required;

                List<dynamic> required = RequiredFields.Where(field => !fields.Contains(field.Slug)).ToList();

                if (required.Count > 0)
                {
                    return String.Join(",", required.Select(d => (string)d.Name).ToList<string>());
                }
            }
            return null;
        }

        private string ValidateUpload(string module, string subModule, List<dynamic> fields)
        {
            dynamic currentModule = UploadFields.FirstOrDefault(d => d.Module == module && d.SubModule == subModule);
            List<dynamic> required = new List<dynamic>();

            if (currentModule != null)
            {
                // Validate always-required fields
                List<dynamic> RequiredFields = currentModule.Required;
                required.AddRange(RequiredFields.Where(field => !fields.Contains(field.Slug)));

                // Validate conditional fields if defined
                if (HasProperty(currentModule, "ConditionalRequired"))
                {
                    foreach (var condition in currentModule.ConditionalRequired)
                    {
                        string conditionSlug = condition.If;
                        if (fields.Contains(conditionSlug))
                        {
                            foreach (var extra in condition.Then)
                            {
                                if (!fields.Contains(extra.Slug))
                                {
                                    required.Add(extra);
                                }
                            }
                        }
                    }
                }
            }

            return required.Count > 0
                ? string.Join(", ", required.Select(d => (string)d.Name))
                : null;
        }

        private void ShowHideButton_Click(object sender, RibbonControlEventArgs e)
        {
            //var currentId = ExcelService.GetStoreDataInWoorkbook("tenantId");

            Globals.ThisAddIn.ToggleTaskPane();
        }

        private void LogoutButton_Click(object sender, RibbonControlEventArgs e)
        {
            User CurrentUser = AuthService.GetStoredUser();

            if (CurrentUser != null)
            {
                AuthService.Logout();
                Globals.ThisAddIn.InitializeAuth();
                MessageBox.Show("You are logged out from ONE plugin");
            }
            else
            {
                Globals.ThisAddIn.ToggleTaskPane();
            }
        }

        private async void UploadButton_Click(object sender, RibbonControlEventArgs e)
        {
            Cursor.Current = Cursors.WaitCursor;

            Microsoft.Office.Interop.Excel.Application excelApp = Globals.ThisAddIn.Application;
            Worksheet activeSheet = excelApp.ActiveSheet as Worksheet;

            Range usedRange = activeSheet.UsedRange;

            int actualRowCount = usedRange.Rows.Count;
            int actualColCount = usedRange.Columns.Count;

            int firstUsedRow = usedRange.Row;
            int firstUsedCol = usedRange.Column;

            // If you need the last used row/column number in the worksheet:
            int lastUsedRow = firstUsedRow + actualRowCount - 1;
            int lastUsedCol = firstUsedCol + actualColCount - 1;

            string fundName = GetFundName(activeSheet, lastUsedRow, lastUsedCol);
            string language = Home.Instance.IsEnglish ? "EN" : "FR";

            if (fundName != null && !String.IsNullOrEmpty(fundName))
            {
                DateTime? navDate = GetNAVDate(activeSheet, lastUsedRow, lastUsedCol);
                if (navDate != null)
                {
                    for (int i = 1; i <= lastUsedRow; i++)
                    {
                        bool hasTable = false;
                        string module = null;
                        string subModule = null;
                        List<dynamic> columns = new List<dynamic>();

                        for (int j = 1; j <= lastUsedCol; j++)
                        {
                            Range currentRange = activeSheet.Cells[i, j] as Range;

                            string cellComment = currentRange.Comment != null ? currentRange.Comment.Text() : null;
                            string cellValue = currentRange.Value2 as string;
                            if (cellComment != null && !String.IsNullOrEmpty(cellComment))
                            {
                                List<string> commentData = cellComment.Split(':').ToList();
                                if (commentData.Count >= 5)
                                {
                                    if (commentData[0] == "U" && commentData[4] == "list")
                                    {
                                        if (!hasTable)
                                        {
                                            hasTable = true;
                                        }

                                        if (module == null)
                                        {
                                            module = commentData[1];
                                        }

                                        if (subModule == null)
                                        {
                                            subModule = commentData[2];
                                        }

                                        columns.Add(new { Name = commentData[3], Position = j });
                                    }

                                }
                            }
                        }

                        if (hasTable && columns.Count > 0 && module == "Funds" && subModule == "NAV")
                        {
                            int dataRow = i + 1;
                            List<dynamic> allTableData = new List<dynamic>();
                            for (int k = dataRow; k <= lastUsedRow; k++)
                            {
                                bool isBreak = false;
                                Dictionary<string, dynamic> currentRowData = new Dictionary<string, dynamic>();
                                for (int col = 0; col < columns.Count; col++)
                                {
                                    Range currentRange = activeSheet.Cells[k, columns[col].Position] as Range;
                                    string cellValue = currentRange.Value2 != null ? currentRange.Value2.ToString() : null;

                                    if (col == 0 && (cellValue == null || String.IsNullOrEmpty(cellValue)))
                                    {
                                        isBreak = true;
                                        break;
                                    }

                                    currentRowData.Add(columns[col].Name, cellValue);
                                }

                                if (isBreak) break;

                                allTableData.Add(currentRowData);
                            }

                            await _uploadDataService.UploadData<dynamic>(allTableData, fundName, navDate?.ToString("yyyy-MM-dd"), subModule, module, language);
                        }
                    }
                }
                else
                {
                    for (int i = 1; i <= lastUsedRow; i++)
                    {
                        bool hasTable = false;
                        string module = null;
                        string subModule = null;
                        List<dynamic> columns = new List<dynamic>();

                        for (int j = 1; j <= lastUsedCol; j++)
                        {
                            Range currentRange = activeSheet.Cells[i, j] as Range;

                            string cellComment = currentRange.Comment != null ? currentRange.Comment.Text() : null;
                            string cellValue = currentRange.Value2 as string;
                            if (cellComment != null && !String.IsNullOrEmpty(cellComment))
                            {
                                List<string> commentData = cellComment.Split(':').ToList();
                                if (commentData.Count >= 5)
                                {
                                    if (commentData[0] == "U" && commentData[4] == "list")
                                    {
                                        if (!hasTable)
                                        {
                                            hasTable = true;
                                        }

                                        if (module == null)
                                        {
                                            module = commentData[1];
                                        }

                                        if (subModule == null)
                                        {
                                            subModule = commentData[2];
                                        }

                                        columns.Add(new { Name = commentData[3], Position = j });
                                    }

                                }
                            }
                        }

                        if (hasTable && columns.Count > 0 && module == "Funds" && subModule == "Accounting")
                        {
                            int dataRow = i + 1;
                            List<dynamic> allTableData = new List<dynamic>();
                            for (int k = dataRow; k <= lastUsedRow; k++)
                            {
                                bool isBreak = false;
                                Dictionary<string, dynamic> currentRowData = new Dictionary<string, dynamic>();
                                for (int col = 0; col < columns.Count; col++)
                                {
                                    Range currentRange = activeSheet.Cells[k, columns[col].Position] as Range;
                                    string cellValue = currentRange.Value2 != null ? currentRange.Value2.ToString() : null;

                                    if (col == 0 && (cellValue == null || String.IsNullOrEmpty(cellValue)))
                                    {
                                        isBreak = true;
                                        break;
                                    }

                                    currentRowData.Add(columns[col].Name, cellValue);
                                }

                                if (isBreak) break;

                                allTableData.Add(currentRowData);
                            }

                            try
                            {
                                await _uploadDataService.UploadData<dynamic>(allTableData, fundName, "", subModule, module, language);
                            }
                            catch (HttpRequestException)
                            {
                                AuthService.Logout();
                                Globals.ThisAddIn.InitializeAuth();
                            }
                        }
                    }
                }

                //string distributionDate = GetDistributionDate(activeSheet, lastUsedRow, lastUsedCol);
                //if(distributionDate != null && !String.IsNullOrEmpty(distributionDate))
                //{

                //}
            }
            else
            {
                for (int i = 1; i <= lastUsedRow; i++)
                {
                    bool hasTable = false;
                    string module = null;
                    string subModule = null;
                    List<dynamic> columns = new List<dynamic>();

                    for (int j = 1; j <= lastUsedCol; j++)
                    {
                        Range currentRange = activeSheet.Cells[i, j] as Range;

                        string cellComment = currentRange.Comment != null ? currentRange.Comment.Text() : null;
                        string cellValue = currentRange.Value2 as string;
                        if (cellComment != null && !String.IsNullOrEmpty(cellComment))
                        {
                            List<string> commentData = cellComment.Split(':').ToList();
                            if (commentData.Count >= 5)
                            {
                                if (commentData[0] == "U" && commentData[4] == "list")
                                {
                                    if (!hasTable)
                                    {
                                        hasTable = true;
                                    }

                                    if (module == null)
                                    {
                                        module = commentData[1];
                                    }

                                    if (subModule == null)
                                    {
                                        subModule = commentData[2];
                                    }

                                    columns.Add(new { Name = commentData[3], Position = j });
                                }

                            }
                        }
                    }

                    if (hasTable && columns.Count > 0)
                    {
                        string errorMessage = ValidateUpload(module, subModule, columns.Select(col => col.Name).ToList());
                        if (errorMessage == null)
                        {
                            MessageBox.Show(
                                  "The upload has started",
                                  "Uploading",
                                  MessageBoxButtons.OK,
                                  MessageBoxIcon.Information
                              );

                            int dataRow = i + 1;
                            List<dynamic> allTableData = new List<dynamic>();
                            for (int k = dataRow; k <= lastUsedRow; k++)
                            {
                                bool isBreak = false;
                                Dictionary<string, dynamic> currentRowData = new Dictionary<string, dynamic>();
                                for (int col = 0; col < columns.Count; col++)
                                {
                                    Range currentRange = activeSheet.Cells[k, columns[col].Position] as Range;

                                    object rawValue = currentRange.Value2;
                                    string cellValue = null;

                                    if (rawValue is double numericValue)
                                    {
                                        // Convert the value to a string with a '.' decimal separator, regardless of system locale
                                        cellValue = numericValue.ToString(CultureInfo.InvariantCulture);
                                    }
                                    else
                                    {
                                        cellValue = currentRange.Value2 != null ? currentRange.Value2.ToString() : null;
                                    }

                                    if (col == 0 && (cellValue == null || String.IsNullOrEmpty(cellValue)))
                                    {
                                        isBreak = true;
                                        break;
                                    }

                                    currentRowData.Add(columns[col].Name, cellValue);
                                }

                                if (isBreak) break;

                                allTableData.Add(currentRowData);
                            }

                            await _uploadDataService.UploadData<dynamic>(allTableData, "", "", subModule, module, language);
                        }
                        else
                        {
                            MessageBox.Show(
                                  $"Please Insert all the mandatory fields before uploading ({errorMessage})",
                                  "Required Fields",
                                  MessageBoxButtons.OK,
                                  MessageBoxIcon.Error
                              );
                        }

                    }
                }
            }

            Cursor.Current = Cursors.Default;
        }

        private string GetFundName(Worksheet activeSheet, int rows, int columns)
        {
            for (int i = 1; i <= 10; i++)
            {
                for (int j = 1; j <= columns; j++)
                {
                    Range currentRange = activeSheet.Cells[i, j] as Range;
                    string cellComment = currentRange.Comment != null ? currentRange.Comment.Text() : null;
                    if (cellComment != null && cellComment.Contains("Funds:Funds:fund-name:value"))
                    {
                        string cellValue = currentRange.Value2 as string;
                        return cellValue;
                    }
                }
            }

            return null;
        }

        private DateTime? GetNAVDate(Worksheet activeSheet, int rows, int columns)
        {
            for (int i = 1; i <= rows; i++)
            {
                for (int j = 1; j <= columns; j++)
                {
                    Range currentRange = activeSheet.Cells[i, j] as Range;
                    string cellComment = currentRange.Comment != null ? currentRange.Comment.Text() : null;
                    if (cellComment != null && cellComment.Contains("Funds:NAV:date:value"))
                    {
                        object cellValue = currentRange.Value;
                        if (cellValue is DateTime)
                        {
                            return (DateTime)cellValue;
                        }
                        else if (cellValue is string)
                        {
                            string dateString = (string)cellValue;
                            
                            // Try MM/DD/YYYY format first
                            if (DateTime.TryParseExact(dateString, "MM/dd/yyyy", null, System.Globalization.DateTimeStyles.None, out DateTime parsedDate))
                            {
                                return parsedDate;
                            }
                            // If MM/DD/YYYY fails, try DD/MM/YYYY format
                            else if (DateTime.TryParseExact(dateString, "dd/MM/yyyy", null, System.Globalization.DateTimeStyles.None, out parsedDate))
                            {
                                return parsedDate;
                            }
                        }
                    }
                }
            }
            return null;
        }

        private string GetDistributionDate(Worksheet activeSheet, int rows, int columns)
        {
            for (int i = 1; i <= rows; i++)
            {
                for (int j = 1; j <= columns; j++)
                {
                    Range currentRange = activeSheet.Cells[i, j] as Range;
                    string cellComment = currentRange.Comment != null ? currentRange.Comment.Text() : null;
                    if (cellComment != null && cellComment.Contains("Limited Partners:Distributions:date:value"))
                    {
                        string cellValue = currentRange.Value2 as string;
                        return cellValue;
                    }
                }
            }
            return null;
        }

        private async void SyncButton_Click(object sender, RibbonControlEventArgs e)
        {
            Microsoft.Office.Interop.Excel.Application excelApp = Globals.ThisAddIn.Application;
            excelApp.ActiveWorkbook.Save();
            Workbook Wb = excelApp.ActiveWorkbook;
            string currentWorkbookName = Wb.Name;
            List<string> nameSegments = currentWorkbookName.Split('-').ToList();
            if (nameSegments.Count >= 4 && nameSegments[0] == "id")
            {
                string currentTemplateFileId = nameSegments[1];
                string fileType = nameSegments[2];
                string currentFileName = nameSegments[nameSegments.Count - 1];

                string tempFilePath = Path.GetTempFileName();
                Wb.SaveCopyAs(tempFilePath);

                try
                {
                    byte[] workbookData;
                    using (FileStream fileStream = new FileStream(tempFilePath, FileMode.Open, FileAccess.Read))
                    {
                        MemoryStream memoryStream = new MemoryStream();
                        fileStream.CopyTo(memoryStream);

                        workbookData = memoryStream.ToArray();
                    }
                    var result = await FileVersionService.UploadExcelFile(workbookData, currentTemplateFileId, currentFileName, fileType);
                    MessageBox.Show("Version Published Successfully!");
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
            }
        }

    }
}
