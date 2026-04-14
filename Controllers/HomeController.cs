using ClosedXML.Excel;
using Microsoft.Extensions.Options;
using Reports.Models;
using Reports.Utilities;
using System.Data.SqlClient;
using System.Diagnostics;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Data;
using Microsoft.AspNetCore.Mvc;
using CrystalDecisions.CrystalReports.Engine;
using CrystalDecisions.Shared;
using Microsoft.Extensions.Hosting.Internal;
using System.IO;
using System.Data.SqlClient;
using DocumentFormat.OpenXml.Math;


namespace Reports.Controllers
{
    public class HomeController : Controller
	{
		private readonly ILogger<HomeController> _logger;
		private readonly IOptions<AppSettings> _options;
		private readonly IWebHostEnvironment _hostingEnvironment;
		public HomeController(ILogger<HomeController> logger, IOptions<AppSettings> options, IWebHostEnvironment hostingEnvironment)
		{
			_logger = logger;
			_options = options;
			_hostingEnvironment = hostingEnvironment;
		}

        public IActionResult Index()
        {
            return View();
        }
        public IActionResult AllMaterialsReport()
		{
			return View(new AllMaterialsReportViewModel());
		}
        
        private Cell CreateCell(string value)
		{
			Cell cell = new Cell(new CellValue(value));
			cell.DataType = new EnumValue<CellValues>(CellValues.String);
			return cell;
		}

        [HttpGet("Home/AllMaterialsReportGetResult")]
        public async Task<IActionResult> AllMaterialsReportGetResult()
        {
            string connectionString = _options.Value.ConnectionString;

            using (SqlConnection connection = new SqlConnection(connectionString))
            {
                await connection.OpenAsync();

                using (SqlCommand command = connection.CreateCommand())
                {
                    command.CommandText = "SELECT [CODE], [DES], [UNT] FROM [dbo].[FMTRL]";

                    using (SqlDataAdapter adapter = new SqlDataAdapter(command))
                    {
                        DataTable dt = new DataTable();
                        adapter.Fill(dt);

                        using (MemoryStream stream = new MemoryStream())
                        {
                            using (SpreadsheetDocument document = SpreadsheetDocument.Create(stream, SpreadsheetDocumentType.Workbook))
                            {
                                WorkbookPart workbookPart = document.AddWorkbookPart();
                                workbookPart.Workbook = new Workbook();

                                WorksheetPart worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
                                worksheetPart.Worksheet = new Worksheet(new SheetData());

                                SheetData sheetData = worksheetPart.Worksheet.GetFirstChild<SheetData>();

                                // Header row
                                Row headerRow = new Row();
                                headerRow.Append(
                                    CreateCell("Code"),
                                    CreateCell("Description"),
                                    CreateCell("Unit Of Measure")
                                );
                                sheetData.Append(headerRow);

                                // Data rows
                                foreach (DataRow row in dt.Rows)
                                {
                                    Row dataRow = new Row();
                                    dataRow.Append(
                                        CreateCell(row["CODE"].ToString()),
                                        CreateCell(row["DES"].ToString()),
                                        CreateCell(row["UNT"].ToString())
                                    );
                                    sheetData.Append(dataRow);
                                }

                                // Add sheet info
                                Sheets sheets = workbookPart.Workbook.AppendChild(new Sheets());
                                Sheet sheet = new Sheet()
                                {
                                    Id = workbookPart.GetIdOfPart(worksheetPart),
                                    SheetId = 1,
                                    Name = "Sheet1"
                                };
                                sheets.Append(sheet);

                                workbookPart.Workbook.Save();
                            }

                            // Reset position before reading
                            stream.Position = 0;

                            // Return the Excel file to browser
                            return File(
                                stream.ToArray(),
                                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                "AllMaterialsReport.xlsx"
                            );
                        }
                    }
                }
            }
        }
        public IActionResult MaterialsDetailsReport()
        {
            return View();
        }
        [Route("/Home/MaterialsDetailsReportGetResult/{year}")]
        public async Task<IActionResult> MaterialsDetailsReportGetResult(string year)
        {
            string connectionString = _options.Value.ConnectionString;
            //string date_str = date.ToString("M/d/yyyy", System.Globalization.CultureInfo.InvariantCulture);
            using (SqlConnection connection = new SqlConnection(connectionString))
            {
                await connection.OpenAsync();

                using (SqlCommand command = connection.CreateCommand())
                {
                    command.CommandText = $"DECLARE @StartYear INT,\r\n\t@MaterialCode CHAR(11) = NULL,\r\n\t @StoreCode INT = NULL\r\nSET NOCOUNT ON\r\n    SET ARITHIGNORE ON\r\n    SET ARITHABORT OFF\r\n    SET ANSI_WARNINGS OFF\r\n    DECLARE @StartOfYear DATETIME;\r\n    DECLARE @EndPrevYear DATETIME;\r\n    DECLARE @CurrentDate DATETIME = GETDATE();\r\n    SET @StartOfYear = DATEFROMPARTS({year}, 1, 1); \r\n    SET @EndPrevYear = DATEADD(ms, -3, @StartOfYear); \r\n    DECLARE @ExcludedStores TABLE (STOR INT PRIMARY KEY);\r\n    INSERT INTO @ExcludedStores (STOR) VALUES (401001),(401002),(401003),(401004),(401005),(401006),(401007),(401008),(401011),(401012),(401017),(401031),(401032),(401036),(401037),(401038),(401039),(401040),(401041),(401042),(402001),(402002),(402003),(402004),(402005),(402006),(402007),(402008),(402011),(402012),(402017),(402031),(402032),(402036),(402037),(402038),(402039),(402040),(402041),(402042),(409001),(409002),(411001),(411002),(411003),(411004),(411005),(419003);\r\n    WITH BeginningBalance AS (\r\n        SELECT\r\n            FINVNT.MTRL,\r\n            SUM(COALESCE(FINVNT.BG_BAL, 0)) + COALESCE(FMVMNT_SUB1.CR_BAL, 0) AS Calculated_BG_BAL\r\n        FROM dbo.FINVNT\r\n        LEFT JOIN \r\n            (\r\n                SELECT\r\n                    FMVMNT.MTRL,\r\n                    COALESCE(SUM(CASE WHEN (FMVMNT.KIND = 2 OR FMVMNT.KIND = 3 OR FMVMNT.KIND = 4) THEN FMVMNT.QUT_UNT ELSE 0 END), 0) -\r\n                    COALESCE(SUM(CASE WHEN (FMVMNT.KIND = 10 OR FMVMNT.KIND = 11) THEN FMVMNT.QUT_UNT ELSE 0 END), 0) AS CR_BAL\r\n                FROM dbo.FMVMNT \r\n                WHERE \r\n                    (FMVMNT.DTE <= @EndPrevYear) \r\n                    AND (@MaterialCode IS NULL OR FMVMNT.MTRL = @MaterialCode)\r\n                    AND (@StoreCode IS NULL OR FMVMNT.STOR = @StoreCode)\r\n                    AND FMVMNT.STOR NOT IN (SELECT STOR FROM @ExcludedStores)\r\n                GROUP BY FMVMNT.MTRL\r\n            ) AS FMVMNT_SUB1 ON FMVMNT_SUB1.MTRL = FINVNT.MTRL\r\n        WHERE\r\n            (@MaterialCode IS NULL OR FINVNT.MTRL = @MaterialCode)\r\n            AND (@StoreCode IS NULL OR FINVNT.STOR = @StoreCode)\r\n            AND FINVNT.STOR NOT IN (SELECT STOR FROM @ExcludedStores)\r\n        GROUP BY FINVNT.MTRL, FMVMNT_SUB1.CR_BAL\r\n    )\r\n    SELECT\r\n\t\tTRIM(M.CODE) AS MTRL_CDE,\r\n\t\tTRIM(M.DES) AS MTRL_DES,\r\n\t\tTRIM(M.UNT) AS UNT,\r\n\t\tCOALESCE(CONVERT(DECIMAL(10, 2), BB.Calculated_BG_BAL), 0) AS BG_BAL, \r\n\t\tCONVERT(DECIMAL(10, 2), \r\n\t\t\tCOALESCE(\r\n\t\t\t\tSUM(CASE WHEN (V.KIND = 2 OR V.KIND = 4) THEN V.QUT_UNT ELSE 0 END), 0)\r\n\t\t) AS QTY_IN, \r\n\t\tCONVERT(DECIMAL(10, 2), \r\n\t\t\tCOALESCE(\r\n\t\t\t\tSUM(CASE WHEN (V.KIND = 10 OR V.KIND = 20 OR V.KIND = 30) THEN V.QUT_UNT ELSE 0 END), 0)\r\n\t\t) AS QTY_OUT, \r\n\t\tCONVERT(DECIMAL(10, 2), \r\n\t\t\tCOALESCE(BB.Calculated_BG_BAL, 0) + \r\n\t\t\tCOALESCE(\r\n\t\t\t\tSUM(CASE WHEN (V.KIND = 2 OR V.KIND = 3 OR V.KIND = 4) THEN V.QUT_UNT ELSE 0 END), 0) -\r\n\t\t\tCOALESCE(\r\n\t\t\t\tSUM(CASE WHEN (V.KIND = 10 OR V.KIND = 11 OR V.KIND = 20 OR V.KIND = 30) THEN V.QUT_UNT ELSE 0 END), 0)\r\n\t\t) AS CURRENT_BAL,\r\n\r\n    -- ✅ Last date for KIND = 4\r\n    MAX(CASE WHEN V.KIND = 4 THEN V.DTE END) AS LAST_PURCHASE_DATE,\r\n\r\n    -- ✅ Last date for KIND = 10\r\n    MAX(CASE WHEN V.KIND = 10 THEN V.DTE END) AS LAST_ISSUE_DATE\r\n\tFROM dbo.FMTRL M \r\n    LEFT JOIN BeginningBalance BB ON M.CODE = BB.MTRL\r\n    LEFT JOIN \r\n        dbo.FMVMNT V ON M.CODE = V.MTRL\r\n        AND V.DTE >= @StartOfYear \r\n        AND V.DTE <= @CurrentDate\r\n        AND (@StoreCode IS NULL OR V.STOR = @StoreCode)\r\n        AND V.STOR NOT IN (SELECT STOR FROM @ExcludedStores)\r\n    WHERE @MaterialCode IS NULL OR M.CODE = @MaterialCode\r\n    GROUP BY M.CODE, M.DES, M.UNT, BB.Calculated_BG_BAL\r\n    ORDER BY M.CODE";

                    using (SqlDataAdapter adapter = new SqlDataAdapter(command))
                    {
                        DataTable dt = new DataTable();
                        adapter.Fill(dt);

                        using (MemoryStream stream = new MemoryStream())
                        {
                            using (SpreadsheetDocument document = SpreadsheetDocument.Create(stream, SpreadsheetDocumentType.Workbook))
                            {
                                WorkbookPart workbookPart = document.AddWorkbookPart();
                                workbookPart.Workbook = new Workbook();

                                WorksheetPart worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
                                worksheetPart.Worksheet = new Worksheet(new SheetData());

                                SheetData sheetData = worksheetPart.Worksheet.GetFirstChild<SheetData>();

                                // Header row
                                Row headerRow = new Row();
                                headerRow.Append(
                                    CreateCell("كود الصنف"),
                                    CreateCell("المواصفة"),
                                    CreateCell("وحدة القياس"),
                                    CreateCell("الرصيد الافتتاحي"),
                                    CreateCell("اجمالي الاضافات"),
                                    CreateCell("اجمالي الصرف"),
                                    CreateCell("الرصيد الاجمالي"),
                                    CreateCell("تاريخ آخر اضافة"),
                                    CreateCell("تاريخ آخر صرف")
                                );
                                sheetData.Append(headerRow);

                                // Data rows
                                foreach (DataRow row in dt.Rows)
                                {
                                    Row dataRow = new Row();
                                    dataRow.Append(
                                        CreateCell(row["MTRL_CDE"].ToString()),
                                        CreateCell(row["MTRL_DES"].ToString()),
                                        CreateCell(row["UNT"].ToString()),
                                        CreateCell(row["BG_BAL"].ToString()),
                                        CreateCell(row["QTY_IN"].ToString()),
                                        CreateCell(row["QTY_OUT"].ToString()),
                                        CreateCell(row["CURRENT_BAL"].ToString()),
                                        CreateCell(
                                        row["LAST_PURCHASE_DATE"] != DBNull.Value
                                            ? Convert.ToDateTime(row["LAST_PURCHASE_DATE"]).ToString("yyyy-MM-dd")
                                            : ""
                                        ),

                                       CreateCell(
                                        row["LAST_ISSUE_DATE"] != DBNull.Value
                                            ? Convert.ToDateTime(row["LAST_ISSUE_DATE"]).ToString("yyyy-MM-dd")
                                            : ""
                                        )
                                    );
                                    sheetData.Append(dataRow);
                                }

                                // Add sheet info
                                Sheets sheets = workbookPart.Workbook.AppendChild(new Sheets());
                                Sheet sheet = new Sheet()
                                {
                                    Id = workbookPart.GetIdOfPart(worksheetPart),
                                    SheetId = 1,
                                    Name = "Sheet1"
                                };
                                sheets.Append(sheet);

                                workbookPart.Workbook.Save();
                            }

                            // Reset position before reading
                            stream.Position = 0;

                            // Return the Excel file to browser
                            return File(
                                stream.ToArray(),
                                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                "تقرير تفاصيل المهمات مستعمل.xlsx"
                            );
                        }
                    }
                }
            }
        }
        public IActionResult MaterialsLocalDetailsReport()
        {
            return View();
        }
        [Route("/Home/MaterialsLocalDetailsReportGetResult/{year}")]
        public async Task<IActionResult> MaterialsLocalDetailsReportGetResult(string year)
        {
            string connectionString = _options.Value.ConnectionString;
            //string date_str = date.ToString("M/d/yyyy", System.Globalization.CultureInfo.InvariantCulture);
            using (SqlConnection connection = new SqlConnection(connectionString))
            {
                await connection.OpenAsync();

                using (SqlCommand command = connection.CreateCommand())
                {
                    command.CommandText = $"DECLARE @StartYear INT,\r\n\t@MaterialCode CHAR(11) = NULL,\r\n\t @StoreCode INT = NULL\r\nSET NOCOUNT ON\r\n    SET ARITHIGNORE ON\r\n    SET ARITHABORT OFF\r\n    SET ANSI_WARNINGS OFF\r\n    DECLARE @StartOfYear DATETIME;\r\n    DECLARE @EndPrevYear DATETIME;\r\n    DECLARE @CurrentDate DATETIME = GETDATE();\r\n    SET @StartOfYear = DATEFROMPARTS({year}, 1, 1); \r\n    SET @EndPrevYear = DATEADD(ms, -3, @StartOfYear); \r\n    DECLARE @ExcludedStores TABLE (STOR INT PRIMARY KEY);\r\n    INSERT INTO @ExcludedStores (STOR) VALUES (403001),(403002),(403003),(403004),(403005),(403006),(403007),(403008),(403011),(403012),(403017),(403031),(431032),(403036),(403037),(403038),(403039),(403040),(403041),(403042),(405001),(405002),(405003),(405004),(405005),(405006),(405007),(405008),(4054011),(405012),(405017),(405031),(405032),(405036),(405037),(405038),(405039),(405040),(405041),(405042),(411001),(411002),(411003),(411004),(411005),(419003);\r\n    WITH BeginningBalance AS (\r\n        SELECT\r\n            FINVNT.MTRL,\r\n            SUM(COALESCE(FINVNT.BG_BAL, 0)) + COALESCE(FMVMNT_SUB1.CR_BAL, 0) AS Calculated_BG_BAL\r\n        FROM dbo.FINVNT\r\n        LEFT JOIN \r\n            (\r\n                SELECT\r\n                    FMVMNT.MTRL,\r\n                    COALESCE(SUM(CASE WHEN (FMVMNT.KIND = 2 OR FMVMNT.KIND = 3 OR FMVMNT.KIND = 4) THEN FMVMNT.QUT_UNT ELSE 0 END), 0) -\r\n                    COALESCE(SUM(CASE WHEN (FMVMNT.KIND = 10 OR FMVMNT.KIND = 11) THEN FMVMNT.QUT_UNT ELSE 0 END), 0) AS CR_BAL\r\n                FROM dbo.FMVMNT \r\n                WHERE \r\n                    (FMVMNT.DTE <= @EndPrevYear) \r\n                    AND (@MaterialCode IS NULL OR FMVMNT.MTRL = @MaterialCode)\r\n                    AND (@StoreCode IS NULL OR FMVMNT.STOR = @StoreCode)\r\n                    AND FMVMNT.STOR NOT IN (SELECT STOR FROM @ExcludedStores)\r\n                GROUP BY FMVMNT.MTRL\r\n            ) AS FMVMNT_SUB1 ON FMVMNT_SUB1.MTRL = FINVNT.MTRL\r\n        WHERE\r\n            (@MaterialCode IS NULL OR FINVNT.MTRL = @MaterialCode)\r\n            AND (@StoreCode IS NULL OR FINVNT.STOR = @StoreCode)\r\n            AND FINVNT.STOR NOT IN (SELECT STOR FROM @ExcludedStores)\r\n        GROUP BY FINVNT.MTRL, FMVMNT_SUB1.CR_BAL\r\n    )\r\n    SELECT\r\n\t\tTRIM(M.CODE) AS MTRL_CDE,\r\n\t\tTRIM(M.DES) AS MTRL_DES,\r\n\t\tTRIM(M.UNT) AS UNT,\r\n\t\tCOALESCE(CONVERT(DECIMAL(10, 2), BB.Calculated_BG_BAL), 0) AS BG_BAL, \r\n\t\tCONVERT(DECIMAL(10, 2), \r\n\t\t\tCOALESCE(\r\n\t\t\t\tSUM(CASE WHEN (V.KIND = 2 OR V.KIND = 4) THEN V.QUT_UNT ELSE 0 END), 0)\r\n\t\t) AS QTY_IN, \r\n\t\tCONVERT(DECIMAL(10, 2), \r\n\t\t\tCOALESCE(\r\n\t\t\t\tSUM(CASE WHEN (V.KIND = 10 OR V.KIND = 20 OR V.KIND = 30) THEN V.QUT_UNT ELSE 0 END), 0)\r\n\t\t) AS QTY_OUT, \r\n\t\tCONVERT(DECIMAL(10, 2), \r\n\t\t\tCOALESCE(BB.Calculated_BG_BAL, 0) + \r\n\t\t\tCOALESCE(\r\n\t\t\t\tSUM(CASE WHEN (V.KIND = 2 OR V.KIND = 3 OR V.KIND = 4) THEN V.QUT_UNT ELSE 0 END), 0) -\r\n\t\t\tCOALESCE(\r\n\t\t\t\tSUM(CASE WHEN (V.KIND = 10 OR V.KIND = 11 OR V.KIND = 20 OR V.KIND = 30) THEN V.QUT_UNT ELSE 0 END), 0)\r\n\t\t) AS CURRENT_BAL,\r\n\r\n    -- ✅ Last date for KIND = 4\r\n    MAX(CASE WHEN V.KIND = 4 THEN V.DTE END) AS LAST_PURCHASE_DATE,\r\n\r\n    -- ✅ Last date for KIND = 10\r\n    MAX(CASE WHEN V.KIND = 10 THEN V.DTE END) AS LAST_ISSUE_DATE\r\n\tFROM dbo.FMTRL M \r\n    LEFT JOIN BeginningBalance BB ON M.CODE = BB.MTRL\r\n    LEFT JOIN \r\n        dbo.FMVMNT V ON M.CODE = V.MTRL\r\n        AND V.DTE >= @StartOfYear \r\n        AND V.DTE <= @CurrentDate\r\n        AND (@StoreCode IS NULL OR V.STOR = @StoreCode)\r\n        AND V.STOR NOT IN (SELECT STOR FROM @ExcludedStores)\r\n    WHERE @MaterialCode IS NULL OR M.CODE = @MaterialCode\r\n    GROUP BY M.CODE, M.DES, M.UNT, BB.Calculated_BG_BAL\r\n    ORDER BY M.CODE";

                    using (SqlDataAdapter adapter = new SqlDataAdapter(command))
                    {
                        DataTable dt = new DataTable();
                        adapter.Fill(dt);

                        using (MemoryStream stream = new MemoryStream())
                        {
                            using (SpreadsheetDocument document = SpreadsheetDocument.Create(stream, SpreadsheetDocumentType.Workbook))
                            {
                                WorkbookPart workbookPart = document.AddWorkbookPart();
                                workbookPart.Workbook = new Workbook();

                                WorksheetPart worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
                                worksheetPart.Worksheet = new Worksheet(new SheetData());

                                SheetData sheetData = worksheetPart.Worksheet.GetFirstChild<SheetData>();

                                // Header row
                                Row headerRow = new Row();
                                headerRow.Append(
                                    CreateCell("كود الصنف"),
                                    CreateCell("المواصفة"),
                                    CreateCell("وحدة القياس"),
                                    CreateCell("الرصيد الافتتاحي"),
                                    CreateCell("اجمالي الاضافات"),
                                    CreateCell("اجمالي الصرف"),
                                    CreateCell("الرصيد الاجمالي"),
                                    CreateCell("تاريخ آخر اضافة"),
                                    CreateCell("تاريخ آخر صرف")
                                );
                                sheetData.Append(headerRow);

                                // Data rows
                                foreach (DataRow row in dt.Rows)
                                {
                                    Row dataRow = new Row();
                                    dataRow.Append(
                                        CreateCell(row["MTRL_CDE"].ToString()),
                                        CreateCell(row["MTRL_DES"].ToString()),
                                        CreateCell(row["UNT"].ToString()),
                                        CreateCell(row["BG_BAL"].ToString()),
                                        CreateCell(row["QTY_IN"].ToString()),
                                        CreateCell(row["QTY_OUT"].ToString()),
                                        CreateCell(row["CURRENT_BAL"].ToString()),
                                        CreateCell(
                                        row["LAST_PURCHASE_DATE"] != DBNull.Value
                                            ? Convert.ToDateTime(row["LAST_PURCHASE_DATE"]).ToString("yyyy-MM-dd")
                                            : ""
                                        ),

                                       CreateCell(
                                        row["LAST_ISSUE_DATE"] != DBNull.Value
                                            ? Convert.ToDateTime(row["LAST_ISSUE_DATE"]).ToString("yyyy-MM-dd")
                                            : ""
                                        )
                                    );
                                    sheetData.Append(dataRow);
                                }

                                // Add sheet info
                                Sheets sheets = workbookPart.Workbook.AppendChild(new Sheets());
                                Sheet sheet = new Sheet()
                                {
                                    Id = workbookPart.GetIdOfPart(worksheetPart),
                                    SheetId = 1,
                                    Name = "Sheet1"
                                };
                                sheets.Append(sheet);

                                workbookPart.Workbook.Save();
                            }

                            // Reset position before reading
                            stream.Position = 0;

                            // Return the Excel file to browser
                            return File(
                                stream.ToArray(),
                                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                "تقرير تفاصيل المهمات جديد.xlsx"
                            );
                        }
                    }
                }
            }
        }
        public IActionResult MaterialsQuantityReport()
		{
			return View();
		}
        [HttpGet]
        public async Task<IActionResult> MaterialsQuantityReportGetResult(string id)
        {
            //string connectionString = @"Data Source=PMS-MB-STOREAPP\MSSQLSERVER,1433;Initial Catalog=PLANDB;User ID=pms\av;Password=K@$perAV;";
            //string connectionString = @"Data Source=.\MSSQLSERVER;Initial Catalog=Test;Trusted_Connection = True;";
            DateOnly date = DateOnly.Parse(id);
            string connectionString = _options.Value.ConnectionString;
            // Convert to string in m/d/yyyy format
            string date_str = date.ToString("M/d/yyyy", System.Globalization.CultureInfo.InvariantCulture);
			//string date_str = date.ToString("M/d/yyyy", System.Globalization.CultureInfo.InvariantCulture);

			using (SqlConnection connection = new SqlConnection(connectionString))
            {
                await connection.OpenAsync();

                // Execute your SQL query using the connection
                using (SqlCommand command = connection.CreateCommand())
                {
                    command.CommandText = $"DECLARE @S_FROM_DTE   DATETIME \r\n DECLARE @S_TO_DTE     DATETIME \r\n DECLARE @S_FROM_CODE  CHAR(11)\r\n DECLARE @S_TO_CODE    CHAR(11)\r\n DECLARE @S_SW         INT\r\n-- \r\n-- \r\n SET @S_FROM_DTE  = '1/1/2024'\r\n SET @S_TO_DTE    = '{date_str}'\r\n SET @S_FROM_CODE = '00000000000'\r\n SET @S_TO_CODE   = 'zzzzzzzzzzz'\r\n SET @S_SW        = 1\r\n\r\n\r\nSET NOCOUNT ON\r\nSET ARITHIGNORE ON\r\nSET ARITHABORT  OFF\r\nSET ANSI_WARNINGS OFF\r\n---------------------------\r\nIF @S_SW = 1 \r\nSELECT \tFINVNT.STOR  ,  \r\n\tFSTOR.DES    ,\r\n\tFINVNT.MTRL  ,\r\n\tFMTRL.DES    AS MTRL_DES ,\r\n\tFMTRL.UNT    ,\t\r\n\tCOALESCE(FMVMNT_SUB.BG_BAL , 0)  + COALESCE(FINVNT.BG_BAL , 0)  +\r\n\tCOALESCE(FMVMNT_SUB1.QUT_IN     , 0)  +\r\n\tCOALESCE(FMVMNT_SUB1.QUT_RET    , 0)  +\r\n\tCOALESCE(FMVMNT_SUB1.QUT_EXIN   , 0)  -\r\n\tCOALESCE(FMVMNT_SUB1.QUT_EXOUT  , 0)  -\r\n\tCOALESCE(FMVMNT_SUB1.QUT_OUT    , 0)  AS RESULT\r\n\t  \r\n\r\n\r\nFROM  FINVNT LEFT JOIN FSTOR ON  FINVNT.STOR = FSTOR.CODE  \r\n\t     LEFT JOIN FMTRL ON  FINVNT.MTRL = FMTRL.CODE  \r\n\r\n--الرصيد الافتتاحى من ملف FMVMNT\r\n--------------------------------\r\n\t     LEFT JOIN (SELECT \tFMVMNT.STOR , FMVMNT.MTRL ,\r\n\t\t\t\tCOALESCE(SUM(CASE WHEN (FMVMNT.KIND = 2 OR\r\n \t\t\t\t\t\t\tFMVMNT.KIND = 3 OR\r\n\t\t\t\t\t\t\tFMVMNT.KIND = 4) \r\n\t\t\t\t\t\t  \tTHEN FMVMNT.QUT_UNT  ELSE 0 END) , 0) -\r\n\t\t\t\tCOALESCE(SUM(CASE WHEN (FMVMNT.KIND = 10 OR\r\n\t\t\t\t\t\t\tFMVMNT.KIND = 11) \r\n\t\t\t\t\t\t  \tTHEN FMVMNT.QUT_UNT  ELSE 0 END) , 0) AS BG_BAL ,\r\n\r\n\t\t\t\t(COALESCE(SUM(CASE WHEN  FMVMNT.KIND = 4 THEN FMVMNT.PRICE  ELSE 0 END) , 0) +\r\n\t\t\t\t COALESCE(SUM(CASE WHEN (FMVMNT.KIND = 2 OR\r\n \t\t\t\t\t\t\t FMVMNT.KIND = 3)\r\n \t\t\t\t\t  \t  THEN (FMVMNT.QUT_UNT * FMVMNT.PRICE)  ELSE 0 END) , 0)) -\r\n\t\t\t\tCOALESCE(SUM(CASE WHEN (FMVMNT.KIND = 10 OR\r\n\t\t\t\t\t\t\tFMVMNT.KIND = 11) \r\n\t\t\t\t\t\t  \tTHEN (FMVMNT.QUT_UNT * FMVMNT.PRICE) ELSE 0 END) , 0) AS BG_VAL\r\n\r\n\r\n\t\t\tFROM  FMVMNT \r\n\t\t\tWHERE (FMVMNT.DTE < @S_FROM_DTE)\r\n\t\t\tGROUP BY FMVMNT.STOR , FMVMNT.MTRL) AS FMVMNT_SUB ON \r\n\t\t\t\t\t\t\t\t FMVMNT_SUB.STOR = FINVNT.STOR AND\r\n\t\t\t\t\t\t\t\t FMVMNT_SUB.MTRL = FINVNT.MTRL\r\n\r\n\r\n--تجميع الحركة من ملف FMVMNT\r\n----------------------------\r\n\t     LEFT JOIN (SELECT \tFMVMNT.STOR , \r\n\t\t\t\tFMVMNT.MTRL  ,\r\n\t\t\t\tCOALESCE(SUM(CASE WHEN (FMVMNT.KIND = 4) THEN FMVMNT.QUT_UNT  ELSE 0 END) , 0) AS QUT_IN     ,\r\n\t\t\t\tCOALESCE(SUM(CASE WHEN (FMVMNT.KIND = 4) THEN (FMVMNT.PRICE)  ELSE 0 END) , 0) AS VAL_IN     ,\r\n\t\t\t\tCOALESCE(SUM(CASE WHEN (FMVMNT.KIND = 2) THEN FMVMNT.QUT_UNT  ELSE 0 END) , 0) AS QUT_RET    ,\r\n\t\t\t\tCOALESCE(SUM(CASE WHEN (FMVMNT.KIND = 2) THEN (FMVMNT.PRICE * FMVMNT.QUT_UNT)  ELSE 0 END) , 0) AS VAL_RET    ,\r\n\t\t\t\tCOALESCE(SUM(CASE WHEN (FMVMNT.KIND = 3) THEN FMVMNT.QUT_UNT  ELSE 0 END) , 0) AS QUT_EXIN   ,\r\n\t\t\t\tCOALESCE(SUM(CASE WHEN (FMVMNT.KIND = 3) THEN (FMVMNT.PRICE * FMVMNT.QUT_UNT)  ELSE 0 END) , 0) AS VAL_EXIN   ,\r\n\t\t\t\tCOALESCE(SUM(CASE WHEN (FMVMNT.KIND = 11)THEN FMVMNT.QUT_UNT  ELSE 0 END) , 0) AS QUT_EXOUT  ,\r\n\t\t\t\tCOALESCE(SUM(CASE WHEN (FMVMNT.KIND = 11)THEN (FMVMNT.PRICE * FMVMNT.QUT_UNT)  ELSE 0 END) , 0) AS VAL_EXOUT  ,\r\n\t\t\t\tCOALESCE(SUM(CASE WHEN (FMVMNT.KIND = 10)THEN FMVMNT.QUT_UNT  ELSE 0 END) , 0) AS QUT_OUT    ,\r\n\t\t\t\tCOALESCE(SUM(CASE WHEN (FMVMNT.KIND = 10)THEN (FMVMNT.PRICE * FMVMNT.QUT_UNT)  ELSE 0 END) , 0) AS VAL_OUT\r\n\r\n\t\t\tFROM  FMVMNT LEFT JOIN FCOST ON FMVMNT.COST = FCOST.CODE\r\n\t\t\tWHERE (FMVMNT.DTE BETWEEN @S_FROM_DTE AND @S_TO_DTE) AND\r\n\t\t\t      (FMVMNT.MTRL BETWEEN @S_FROM_CODE AND @S_TO_CODE)\r\n\r\n\t\t\tGROUP BY FMVMNT.STOR  , \r\n\t\t\t\t FMVMNT.MTRL  ) AS FMVMNT_SUB1 ON FMVMNT_SUB1.STOR = FINVNT.STOR AND\r\n\t\t\t\t\t\t\t\t  FMVMNT_SUB1.MTRL = FINVNT.MTRL\r\n\r\nWHERE (FINVNT.MTRL BETWEEN @S_FROM_CODE AND @S_TO_CODE)\r\n\r\nGROUP BY FINVNT.STOR         ,  \r\n\tFSTOR.DES            ,\r\n\tFINVNT.MTRL          ,\r\n\tFMTRL.DES            ,\r\n\tFMTRL.UNT            ,\r\n\tFMVMNT_SUB.BG_BAL    ,\r\n\tFINVNT.BG_BAL        ,\r\n\tFMVMNT_SUB.BG_VAL    , \r\n\tFINVNT.BG_PRICE      ,\r\n\tFMVMNT_SUB1.QUT_IN     ,\r\n\tFMVMNT_SUB1.QUT_RET    ,\r\n\tFMVMNT_SUB1.QUT_EXIN   ,\r\n\tFMVMNT_SUB1.QUT_EXOUT  ,\r\n\tFMVMNT_SUB1.QUT_OUT    ,\r\n\tFMVMNT_SUB1.VAL_IN     ,\r\n\tFMVMNT_SUB1.VAL_RET    ,\r\n\tFMVMNT_SUB1.VAL_EXIN   ,\r\n\tFMVMNT_SUB1.VAL_EXOUT  ,\r\n\tFMVMNT_SUB1.VAL_OUT    \r\n\r\n\r\nHAVING \r\n\t(COALESCE(FMVMNT_SUB.BG_BAL , 0)  + COALESCE(FINVNT.BG_BAL , 0)  > 0 OR\r\n\tCOALESCE(FMVMNT_SUB1.QUT_IN     , 0)  > 0    OR\r\n\tCOALESCE(FMVMNT_SUB1.QUT_RET    , 0)  > 0    OR\r\n\tCOALESCE(FMVMNT_SUB1.QUT_EXIN   , 0)  > 0    OR\r\n\tCOALESCE(FMVMNT_SUB1.QUT_EXOUT  , 0)  > 0    OR\r\n\tCOALESCE(FMVMNT_SUB1.QUT_OUT    , 0)  > 0    )  \r\n\r\nORDER BY FINVNT.MTRL\r\n\r\n\r\n\r\n\r\nIF @S_SW = 2 \r\nSELECT \tFINVNT.STOR  ,  \r\n\tFSTOR.DES    ,\r\n\tFINVNT.MTRL  ,\r\n\tFMTRL.DES    AS MTRL_DES ,\r\n\tFMTRL.UNT    ,\t\r\n\tCOALESCE(FMVMNT_SUB.BG_BAL , 0)  + COALESCE(FINVNT.BG_BAL , 0)  +\r\n\tCOALESCE(FMVMNT_SUB1.QUT_IN     , 0)  +\r\n\tCOALESCE(FMVMNT_SUB1.QUT_RET    , 0)  +\r\n\tCOALESCE(FMVMNT_SUB1.QUT_EXIN   , 0)  -\r\n\tCOALESCE(FMVMNT_SUB1.QUT_EXOUT  , 0)  -\r\n\tCOALESCE(FMVMNT_SUB1.QUT_OUT    , 0)  AS RESULT \r\n\r\n\r\nFROM  FINVNT LEFT JOIN FSTOR ON  FINVNT.STOR = FSTOR.CODE  \r\n\t     LEFT JOIN FMTRL ON  FINVNT.MTRL = FMTRL.CODE  \r\n\r\n--الرصيد الافتتاحى من ملف FMVMNT\r\n--------------------------------\r\n\t     LEFT JOIN (SELECT \tFMVMNT.STOR , FMVMNT.MTRL ,\r\n\t\t\t\tCOALESCE(SUM(CASE WHEN (FMVMNT.KIND = 2 OR\r\n \t\t\t\t\t\t\tFMVMNT.KIND = 3 OR\r\n\t\t\t\t\t\t\tFMVMNT.KIND = 4) \r\n\t\t\t\t\t\t  \tTHEN FMVMNT.QUT_UNT  ELSE 0 END) , 0) -\r\n\t\t\t\tCOALESCE(SUM(CASE WHEN (FMVMNT.KIND = 10 OR\r\n\t\t\t\t\t\t\tFMVMNT.KIND = 11) \r\n\t\t\t\t\t\t  \tTHEN FMVMNT.QUT_UNT  ELSE 0 END) , 0) AS BG_BAL ,\r\n\r\n\t\t\t\t(COALESCE(SUM(CASE WHEN  FMVMNT.KIND = 4 THEN FMVMNT.PRICE  ELSE 0 END) , 0) +\r\n\t\t\t\t COALESCE(SUM(CASE WHEN (FMVMNT.KIND = 2 OR\r\n \t\t\t\t\t\t\t FMVMNT.KIND = 3)\r\n \t\t\t\t\t  \t  THEN (FMVMNT.QUT_UNT * FMVMNT.PRICE)  ELSE 0 END) , 0)) -\r\n\t\t\t\tCOALESCE(SUM(CASE WHEN (FMVMNT.KIND = 10 OR\r\n\t\t\t\t\t\t\tFMVMNT.KIND = 11) \r\n\t\t\t\t\t\t  \tTHEN (FMVMNT.QUT_UNT * FMVMNT.PRICE) ELSE 0 END) , 0) AS BG_VAL\r\n\r\n\r\n\t\t\tFROM  FMVMNT \r\n\t\t\tWHERE (FMVMNT.DTE < @S_FROM_DTE)\r\n\t\t\tGROUP BY FMVMNT.STOR , FMVMNT.MTRL) AS FMVMNT_SUB ON \r\n\t\t\t\t\t\t\t\t FMVMNT_SUB.STOR = FINVNT.STOR AND\r\n\t\t\t\t\t\t\t\t FMVMNT_SUB.MTRL = FINVNT.MTRL\r\n\r\n\r\n--تجميع الحركة من ملف FMVMNT\r\n----------------------------\r\n\t     LEFT JOIN (SELECT \tFMVMNT.STOR , \r\n\t\t\t\tFMVMNT.MTRL  ,\r\n\t\t\t\tCOALESCE(SUM(CASE WHEN (FMVMNT.KIND = 4) THEN FMVMNT.QUT_UNT  ELSE 0 END) , 0) AS QUT_IN     ,\r\n\t\t\t\tCOALESCE(SUM(CASE WHEN (FMVMNT.KIND = 4) THEN (FMVMNT.PRICE)  ELSE 0 END) , 0) AS VAL_IN     ,\r\n\t\t\t\tCOALESCE(SUM(CASE WHEN (FMVMNT.KIND = 2) THEN FMVMNT.QUT_UNT  ELSE 0 END) , 0) AS QUT_RET    ,\r\n\t\t\t\tCOALESCE(SUM(CASE WHEN (FMVMNT.KIND = 2) THEN (FMVMNT.PRICE * FMVMNT.QUT_UNT)  ELSE 0 END) , 0) AS VAL_RET    ,\r\n\t\t\t\tCOALESCE(SUM(CASE WHEN (FMVMNT.KIND = 3) THEN FMVMNT.QUT_UNT  ELSE 0 END) , 0) AS QUT_EXIN   ,\r\n\t\t\t\tCOALESCE(SUM(CASE WHEN (FMVMNT.KIND = 3) THEN (FMVMNT.PRICE * FMVMNT.QUT_UNT)  ELSE 0 END) , 0) AS VAL_EXIN   ,\r\n\t\t\t\tCOALESCE(SUM(CASE WHEN (FMVMNT.KIND = 11)THEN FMVMNT.QUT_UNT  ELSE 0 END) , 0) AS QUT_EXOUT  ,\r\n\t\t\t\tCOALESCE(SUM(CASE WHEN (FMVMNT.KIND = 11)THEN (FMVMNT.PRICE * FMVMNT.QUT_UNT)  ELSE 0 END) , 0) AS VAL_EXOUT  ,\r\n\t\t\t\tCOALESCE(SUM(CASE WHEN (FMVMNT.KIND = 10)THEN FMVMNT.QUT_UNT  ELSE 0 END) , 0) AS QUT_OUT    ,\r\n\t\t\t\tCOALESCE(SUM(CASE WHEN (FMVMNT.KIND = 10)THEN (FMVMNT.PRICE * FMVMNT.QUT_UNT)  ELSE 0 END) , 0) AS VAL_OUT\r\n\r\n\t\t\tFROM  FMVMNT LEFT JOIN FCOST ON FMVMNT.COST = FCOST.CODE\r\n\t\t\tWHERE (FMVMNT.DTE BETWEEN @S_FROM_DTE AND @S_TO_DTE) AND\r\n\t\t\t      (FMVMNT.MTRL BETWEEN @S_FROM_CODE AND @S_TO_CODE)\r\n\r\n\t\t\tGROUP BY FMVMNT.STOR  , \r\n\t\t\t\t FMVMNT.MTRL  ) AS FMVMNT_SUB1 ON FMVMNT_SUB1.STOR = FINVNT.STOR AND\r\n\t\t\t\t\t\t\t\t  FMVMNT_SUB1.MTRL = FINVNT.MTRL\r\n\r\nWHERE (FINVNT.MTRL BETWEEN @S_FROM_CODE AND @S_TO_CODE)\r\n\r\nGROUP BY FINVNT.STOR         ,  \r\n\tFSTOR.DES            ,\r\n\tFINVNT.MTRL          ,\r\n\tFMTRL.DES            ,\r\n\tFMTRL.UNT            ,\r\n\tFMVMNT_SUB.BG_BAL    ,\r\n\tFINVNT.BG_BAL        ,\r\n\tFMVMNT_SUB.BG_VAL    , \r\n\tFINVNT.BG_PRICE      ,\r\n\tFMVMNT_SUB1.QUT_IN     ,\r\n\tFMVMNT_SUB1.QUT_RET    ,\r\n\tFMVMNT_SUB1.QUT_EXIN   ,\r\n\tFMVMNT_SUB1.QUT_EXOUT  ,\r\n\tFMVMNT_SUB1.QUT_OUT    ,\r\n\tFMVMNT_SUB1.VAL_IN     ,\r\n\tFMVMNT_SUB1.VAL_RET    ,\r\n\tFMVMNT_SUB1.VAL_EXIN   ,\r\n\tFMVMNT_SUB1.VAL_EXOUT  ,\r\n\tFMVMNT_SUB1.VAL_OUT    \r\n\r\n\r\nHAVING \r\n--\t(COALESCE(FMVMNT_SUB.BG_BAL , 0)  + COALESCE(FINVNT.BG_BAL , 0)  > 0 OR\r\n\t(COALESCE(FMVMNT_SUB1.QUT_IN     , 0)  > 0    OR\r\n\tCOALESCE(FMVMNT_SUB1.QUT_RET    , 0)  > 0    OR\r\n\tCOALESCE(FMVMNT_SUB1.QUT_EXIN   , 0)  > 0    OR\r\n\tCOALESCE(FMVMNT_SUB1.QUT_EXOUT  , 0)  > 0    OR\r\n\tCOALESCE(FMVMNT_SUB1.QUT_OUT    , 0)  > 0    )  \r\n\r\nORDER BY FINVNT.MTRL\r\n";
                    //command.CommandText = "SELECT [CODE]\r\n      ,REPLACE(REPLACE([DES], CHAR(13), ''), CHAR(10), '')\r\n      ,[UNT]\r\n  FROM [dbo].[FMTRL]";
                    //command.CommandText = "select * from [dbo].[Categories]";
                    command.CommandTimeout = 300;
                    using (SqlDataAdapter adapter = new SqlDataAdapter(command))
                    {
                        ////////////////

                        DataTable dt = new DataTable();
                        adapter.Fill(dt);
                        #region 4
                        // Create a new Excel document
                        using (MemoryStream stream = new MemoryStream())
                        {
                            using (SpreadsheetDocument document = SpreadsheetDocument.Create(stream, SpreadsheetDocumentType.Workbook))
                            {
                                // Add a workbook part to the document
                                WorkbookPart workbookPart = document.AddWorkbookPart();
                                workbookPart.Workbook = new Workbook();

                                // Add a worksheet part to the workbook
                                WorksheetPart worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
                                worksheetPart.Worksheet = new Worksheet(new SheetData());

                                // Get the sheet data of the worksheet
                                SheetData sheetData = worksheetPart.Worksheet.GetFirstChild<SheetData>();

                                // Add headers to the sheet data
                                Row headerRow = new Row();
                                headerRow.Append(
                                    CreateCell("المخزن"),
                                    CreateCell("اسم المخزن"),
                                    CreateCell("كود الصنف"),
                                    CreateCell("المواصفة"),
                                    CreateCell("الوحدة"),
                                    CreateCell("الرصيد")
                                );
                                sheetData.Append(headerRow);

                                // Add data rows to the sheet data
                                foreach (DataRow row in dt.Rows)
                                {
                                    Row dataRow = new Row();
                                    dataRow.Append(
                                        CreateCell(row["STOR"].ToString()),
                                        CreateCell(row["DES"].ToString()),
                                        CreateCell(row["MTRL"].ToString().Trim()),
                                        CreateCell(row["MTRL_DES"].ToString()),
                                        CreateCell(row["UNT"].ToString()),
                                        CreateCell(Convert.ToDecimal(row["RESULT"]).ToString("G29"))
                                    );
                                    sheetData.Append(dataRow);
                                }

                                // Add the worksheet to the workbook
                                Sheets sheets = workbookPart.Workbook.AppendChild(new Sheets());
                                Sheet sheet = new Sheet() { Id = workbookPart.GetIdOfPart(worksheetPart), SheetId = 1, Name = "Sheet1" };
                                sheets.Append(sheet);

                                // Save the workbook
                                workbookPart.Workbook.Save();

                            }
                            // Reset position before reading
                            stream.Position = 0;

                            // Return the Excel file to browser
                            return File(
                                stream.ToArray(),
                                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                "تقرير ارصدة المهمات.xlsx"
                            );
                        }
                        #endregion 4
                        #region 3
                        //// Create a new Excel package
                        //using (ExcelPackage package = new ExcelPackage())
                        //{							
                        //	// Create the worksheet
                        //	ExcelWorksheet worksheet = package.Workbook.Worksheets.Add("Sheet1");

                        //	// Load data from DataTable to the worksheet
                        //	worksheet.Cells["A1"].LoadFromDataTable(dt, true);

                        //	// Set the content type and file name for the response
                        //	Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
                        //	Response.Headers.Add("content-disposition", "attachment; filename=report.xlsx");

                        //	// Write the Excel package to the response output stream
                        //	Response.Body.WriteAsync(package.GetAsByteArray());
                        //}
                        #endregion 3
                        ////////////////
                        #region 2
                        //using (XLWorkbook wb = new XLWorkbook())
                        //{
                        //	string sheetName = "Sheet1"; // Default sheet name
                        //	if (!string.IsNullOrEmpty(dt.TableName))
                        //	{
                        //		sheetName = dt.TableName; // Use the DataTable name as the sheet name if available
                        //	}
                        //	wb.Worksheets.Add(dt, sheetName);
                        //	using (MemoryStream stream = new MemoryStream())
                        //	{
                        //		wb.SaveAs(stream);
                        //		try
                        //		{
                        //			string userProfileDirectory = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile);
                        //			string filePath = Path.Combine(userProfileDirectory, "Downloads", "result.xlsx");
                        //			File(stream.ToArray(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", filePath);
                        //			//File(stream.ToArray(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "%USERPROFILE%\\Downloads\\result.xlsx");									
                        //		}
                        //		catch (Exception)
                        //		{
                        //			Exception customEx = new Exception("error in downloading");
                        //			throw customEx;
                        //		}								
                        //	}
                        //}
                        #endregion 2
                        #region 1
                        //// Process the query result
                        //var lines = new List<string>();

                        //string[] columnNames = result.Columns
                        //	.Cast<DataColumn>()
                        //	.Select(column => column.ColumnName)
                        //	.ToArray();

                        //var header = string.Join(",", columnNames.Select(name => $"\"{name}\""));
                        //lines.Add(header);

                        //var valueLines = result.AsEnumerable()
                        //	.Select(row => string.Join(",", row.ItemArray.Select(val => $"\"{val}\"")));

                        //lines.AddRange(valueLines);

                        //System.IO.File.WriteAllLines(string.IsNullOrEmpty(path)? "C:\\Result.csv" : path, lines);
                        #endregion 1
                    }
                }
            }
        }

		public IActionResult AssetsQuantityReport()
		{
			return View();
		}
		[Route("/Home/AssetsQuantityReportGetResult/{date}")]
		public IActionResult AssetsQuantityReportGetResult(DateOnly date)
		{
			//string connectionString = @"Data Source=PMS-MB-STOREAPP\MSSQLSERVER,1433;Initial Catalog=PLANDB;User ID=pms\av;Password=K@$perAV;";
			//string connectionString = @"Data Source=.\MSSQLSERVER;Initial Catalog=Test;Trusted_Connection = True;";
			string connectionString = _options.Value.ConnectionString;

			// Convert to string in m/d/yyyy format
			string date_str = date.ToString("M/d/yyyy", System.Globalization.CultureInfo.InvariantCulture);

			using (SqlConnection connection = new SqlConnection(connectionString))
			{
				connection.Open();

				// Execute your SQL query using the connection
				using (SqlCommand command = connection.CreateCommand())
				{
					command.CommandText = $" DECLARE @S_FROM_DTE   DATETIME \r\n DECLARE @S_TO_DTE     DATETIME \r\n \r\n \r\n SET @S_FROM_DTE  = '{date_str}'\r\n SET @S_TO_DTE    = '{date_str}'\r\n-- \r\n\r\nSET NOCOUNT ON\r\nSET ARITHIGNORE ON\r\nSET ARITHABORT  OFF\r\nSET ANSI_WARNINGS OFF\r\n---------------------------\r\nSELECT \tFAINVNT.STOR  ,  \r\n\tFASTOR.DES    ,\r\n\tFAINVNT.MTRL  ,\r\n\tFASSIT.DES    AS ASSIT_DES ,\r\n\tFASSIT.UNT    ,\r\n\tCOALESCE(FAMVMNT_SUB.BG_BAL , 0)  + COALESCE(FAINVNT.BG_BAL , 0)  AS BG_BAL ,\r\n\tCOALESCE(FAMVMNT_SUB1.QUT_IN     , 0)  AS IN_BAL    ,\r\n\tCOALESCE(FAMVMNT_SUB1.QUT_RET    , 0)  AS RET_BAL   ,\r\n\tCOALESCE(FAMVMNT_SUB1.QUT_EXIN   , 0)  AS EXIN_BAL  ,\r\n\tCOALESCE(FAMVMNT_SUB1.QUT_EXOUT  , 0)  AS EXOUT_BAL ,\r\n\tCOALESCE(FAMVMNT_SUB1.QUT_OUT    , 0)  AS OUT_BAL   ,\r\nCOALESCE(FAMVMNT_SUB1.QUT_OUT1    , 0)  AS OUT_BAL1   ,\r\nCOALESCE(FAMVMNT_SUB1.QUT_OUT2    , 0)  AS OUT_BAL2   ,\r\n\r\n\tCOALESCE(FAMVMNT_SUB.BG_BAL , 0)  + COALESCE(FAINVNT.BG_BAL , 0) +\r\n\tCOALESCE(FAMVMNT_SUB1.QUT_IN     , 0)  +\r\n\tCOALESCE(FAMVMNT_SUB1.QUT_RET    , 0)  +\r\n\tCOALESCE(FAMVMNT_SUB1.QUT_EXIN   , 0)  -\r\n\tCOALESCE(FAMVMNT_SUB1.QUT_EXOUT  , 0)  -\r\nCOALESCE(FAMVMNT_SUB1.QUT_OUT1    , 0) -\r\nCOALESCE(FAMVMNT_SUB1.QUT_OUT2    , 0) -\r\n\tCOALESCE(FAMVMNT_SUB1.QUT_OUT    , 0)  AS BAL\r\n\r\n\r\nFROM  FAINVNT LEFT JOIN FASTOR ON  FAINVNT.STOR = FASTOR.CODE  \r\n\t     LEFT JOIN FASSIT ON  FAINVNT.MTRL = FASSIT.CODE  \r\n\r\n--الرصيد الافتتاحى من ملف FAMVMNT\r\n--------------------------------\r\n\t     LEFT JOIN (SELECT \tFAMVMNT.STOR , FAMVMNT.MTRL ,\r\n\t\t\t\tCOALESCE(SUM(CASE WHEN (FAMVMNT.KIND = 2 OR\r\n \t\t\t\t\t\t\tFAMVMNT.KIND = 3 OR\r\n\t\t\t\t\t\t\tFAMVMNT.KIND = 4) \r\n\t\t\t\t\t\t  \tTHEN FAMVMNT.QUT_UNT  ELSE 0 END) , 0) -\r\n\r\n\t\t\t\tCOALESCE(SUM(CASE WHEN (FAMVMNT.KIND = 10 OR\r\n\t\t\t\t\t\t\tFAMVMNT.KIND = 11 OR\r\nFAMVMNT.KIND = 20 OR\r\nFAMVMNT.KIND = 30)\r\n\t\t\t\t\t\t  \tTHEN FAMVMNT.QUT_UNT  ELSE 0 END) , 0) AS BG_BAL\r\n\t\t\tFROM  FAMVMNT \r\n\t\t\tWHERE FAMVMNT.DTE < @S_FROM_DTE\r\n\t\t\tGROUP BY FAMVMNT.STOR , FAMVMNT.MTRL) AS \tFAMVMNT_SUB ON \r\n\t\t\t\t\t\t\t\tFAMVMNT_SUB.STOR = FAINVNT.STOR AND\r\n\t\t\t\t\t\t\t\tFAMVMNT_SUB.MTRL = FAINVNT.MTRL\r\n\r\n\r\n--تجميع الحركة من ملف FAMVMNT\r\n----------------------------\r\n\t     LEFT JOIN (SELECT \tFAMVMNT.STOR , FAMVMNT.MTRL ,\r\n\t\t\t\tCOALESCE(SUM(CASE WHEN (FAMVMNT.KIND = 4) THEN FAMVMNT.QUT_UNT  ELSE 0 END) , 0) AS QUT_IN     ,\r\n\t\t\t\tCOALESCE(SUM(CASE WHEN (FAMVMNT.KIND = 2) THEN FAMVMNT.QUT_UNT  ELSE 0 END) , 0) AS QUT_RET    ,\r\n\t\t\t\tCOALESCE(SUM(CASE WHEN (FAMVMNT.KIND = 3) THEN FAMVMNT.QUT_UNT  ELSE 0 END) , 0) AS QUT_EXIN   ,\r\n\t\t\t\tCOALESCE(SUM(CASE WHEN (FAMVMNT.KIND = 11)THEN FAMVMNT.QUT_UNT  ELSE 0 END) , 0) AS QUT_EXOUT  ,\r\n\t\t\t\tCOALESCE(SUM(CASE WHEN (FAMVMNT.KIND = 10)THEN FAMVMNT.QUT_UNT  ELSE 0 END) , 0) AS QUT_OUT,\r\nCOALESCE(SUM(CASE WHEN (FAMVMNT.KIND = 20)THEN FAMVMNT.QUT_UNT  ELSE 0 END) , 0) AS QUT_OUT1,\r\nCOALESCE(SUM(CASE WHEN (FAMVMNT.KIND = 30)THEN FAMVMNT.QUT_UNT  ELSE 0 END) , 0) AS QUT_OUT2\r\n\t\t\tFROM  FAMVMNT \r\n\t\t\tWHERE (FAMVMNT.DTE BETWEEN @S_FROM_DTE AND @S_TO_DTE)\r\n\t\t\tGROUP BY FAMVMNT.STOR , FAMVMNT.MTRL) AS FAMVMNT_SUB1 ON \r\n\t\t\t\t\t\t\t\tFAMVMNT_SUB1.STOR = FAINVNT.STOR AND\r\n\t\t\t\t\t\t\t\tFAMVMNT_SUB1.MTRL = FAINVNT.MTRL\r\n\r\n\r\nGROUP BY FAINVNT.STOR      ,  \r\n\tFASTOR.DES         ,\r\n\tFAINVNT.MTRL       ,\r\n\tFASSIT.DES         ,\r\n\tFASSIT.UNT         ,\r\n\tFAMVMNT_SUB.BG_BAL ,\r\n\tFAINVNT.BG_BAL     ,\r\n\tFAMVMNT_SUB1.QUT_IN     ,\r\n\tFAMVMNT_SUB1.QUT_RET    ,\r\n\tFAMVMNT_SUB1.QUT_EXIN  ,\r\n\tFAMVMNT_SUB1.QUT_EXOUT ,\r\n\tFAMVMNT_SUB1.QUT_OUT    ,\r\nFAMVMNT_SUB1.QUT_OUT1   ,\r\nFAMVMNT_SUB1.QUT_OUT2\r\n";
					//command.CommandText = "SELECT [CODE]\r\n      ,REPLACE(REPLACE([DES], CHAR(13), ''), CHAR(10), '')\r\n      ,[UNT]\r\n  FROM [dbo].[FMTRL]";
					//command.CommandText = "select * from [dbo].[Categories]";

					using (SqlDataAdapter adapter = new SqlDataAdapter(command))
					{
						////////////////

						DataTable dt = new DataTable();
						adapter.Fill(dt);
						#region 4
						// Create a new Excel document
						using (MemoryStream stream = new MemoryStream())
						{
							using (SpreadsheetDocument document = SpreadsheetDocument.Create(stream, SpreadsheetDocumentType.Workbook))
							{
								// Add a workbook part to the document
								WorkbookPart workbookPart = document.AddWorkbookPart();
								workbookPart.Workbook = new Workbook();

								// Add a worksheet part to the workbook
								WorksheetPart worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
								worksheetPart.Worksheet = new Worksheet(new SheetData());

								// Get the sheet data of the worksheet
								SheetData sheetData = worksheetPart.Worksheet.GetFirstChild<SheetData>();

								// Add headers to the sheet data
								Row headerRow = new Row();
								headerRow.Append(
									CreateCell("المخزن"),
									CreateCell("اسم المخزن"),
									CreateCell("كود الصنف"),
									CreateCell("المواصفة"),
									CreateCell("الوحدة"),
									CreateCell("الرصيد")
								);
								sheetData.Append(headerRow);

								// Add data rows to the sheet data
								foreach (DataRow row in dt.Rows)
								{
									Row dataRow = new Row();
									dataRow.Append(
										CreateCell(row["STOR"].ToString()),
										CreateCell(row["DES"].ToString()),
										CreateCell(row["MTRL"].ToString().Trim()),
										CreateCell(row["ASSIT_DES"].ToString()),
										CreateCell(row["UNT"].ToString()),
										CreateCell(Convert.ToDecimal(row["BAL"]).ToString("G29"))

									);
									sheetData.Append(dataRow);
								}

								// Add the worksheet to the workbook
								Sheets sheets = workbookPart.Workbook.AppendChild(new Sheets());
								Sheet sheet = new Sheet() { Id = workbookPart.GetIdOfPart(worksheetPart), SheetId = 1, Name = "Sheet1" };
								sheets.Append(sheet);

								// Save the workbook
								workbookPart.Workbook.Save();

							}
                            // Reset position before reading
                            stream.Position = 0;

                            // Return the Excel file to browser
                            return File(
                                stream.ToArray(),
                                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                "تقرير ارصدة الاصول.xlsx"
                            );
						}
						#endregion 4
						#region 3
						//// Create a new Excel package
						//using (ExcelPackage package = new ExcelPackage())
						//{							
						//	// Create the worksheet
						//	ExcelWorksheet worksheet = package.Workbook.Worksheets.Add("Sheet1");

						//	// Load data from DataTable to the worksheet
						//	worksheet.Cells["A1"].LoadFromDataTable(dt, true);

						//	// Set the content type and file name for the response
						//	Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
						//	Response.Headers.Add("content-disposition", "attachment; filename=report.xlsx");

						//	// Write the Excel package to the response output stream
						//	Response.Body.WriteAsync(package.GetAsByteArray());
						//}
						#endregion 3
						////////////////
						#region 2
						//using (XLWorkbook wb = new XLWorkbook())
						//{
						//	string sheetName = "Sheet1"; // Default sheet name
						//	if (!string.IsNullOrEmpty(dt.TableName))
						//	{
						//		sheetName = dt.TableName; // Use the DataTable name as the sheet name if available
						//	}
						//	wb.Worksheets.Add(dt, sheetName);
						//	using (MemoryStream stream = new MemoryStream())
						//	{
						//		wb.SaveAs(stream);
						//		try
						//		{
						//			string userProfileDirectory = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile);
						//			string filePath = Path.Combine(userProfileDirectory, "Downloads", "result.xlsx");
						//			File(stream.ToArray(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", filePath);
						//			//File(stream.ToArray(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "%USERPROFILE%\\Downloads\\result.xlsx");									
						//		}
						//		catch (Exception)
						//		{
						//			Exception customEx = new Exception("error in downloading");
						//			throw customEx;
						//		}								
						//	}
						//}
						#endregion 2
						#region 1
						//// Process the query result
						//var lines = new List<string>();

						//string[] columnNames = result.Columns
						//	.Cast<DataColumn>()
						//	.Select(column => column.ColumnName)
						//	.ToArray();

						//var header = string.Join(",", columnNames.Select(name => $"\"{name}\""));
						//lines.Add(header);

						//var valueLines = result.AsEnumerable()
						//	.Select(row => string.Join(",", row.ItemArray.Select(val => $"\"{val}\"")));

						//lines.AddRange(valueLines);

						//System.IO.File.WriteAllLines(string.IsNullOrEmpty(path)? "C:\\Result.csv" : path, lines);
						#endregion 1

						return View("Done");
					}
				}
			}
		}

		public IActionResult PersonalCustodyReport()
		{
			return View(new AllMaterialsReportViewModel());
		}
		[Route("/Home/PersonalCustodyReportGetResult")]
		[Route("/Home/PersonalCustodyReportGetResult/{id}")]		
		public async Task<IActionResult> PersonalCustodyReportGetResult(string id = "")
		{			
			string connectionString = _options.Value.ConnectionString;

			using (SqlConnection connection = new SqlConnection(connectionString))
			{
                await connection.OpenAsync();

                // Execute your SQL query using the connection
                using (SqlCommand command = connection.CreateCommand())
				{					
					if (id == "")
						command.CommandText = $"DECLARE\t@return_value int\r\n\r\nEXEC\t@return_value = [dbo].[ASSTP0060A]\r\n\t\t@S_FROM_EMP = 0 ,\r\n\t\t@S_TO_EMP = 999999,\r\n\t\t@S_FROM_LOC = '',\r\n\t\t@S_TO_LOC = 'ZZZZZZZZZZ'";
					else
						command.CommandText = $"DECLARE\t@return_value int\r\n\r\nEXEC\t@return_value = [dbo].[ASSTP0060A]\r\n\t\t@S_FROM_EMP = {id} ,\r\n\t\t@S_TO_EMP = {id},\r\n\t\t@S_FROM_LOC = '',\r\n\t\t@S_TO_LOC = 'ZZZZZZZZZZ'";
                    command.CommandTimeout = 300;
                    using (SqlDataAdapter adapter = new SqlDataAdapter(command))
					{
						DataTable dt = new DataTable();
						adapter.Fill(dt);
						#region 2
						// Create a new Excel document
						using (MemoryStream stream = new MemoryStream())
						{
							using (SpreadsheetDocument document = SpreadsheetDocument.Create(stream, SpreadsheetDocumentType.Workbook))
							{
								// Add a workbook part to the document
								WorkbookPart workbookPart = document.AddWorkbookPart();
								workbookPart.Workbook = new Workbook();

								// Add a worksheet part to the workbook
								WorksheetPart worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
								worksheetPart.Worksheet = new Worksheet(new SheetData());

								// Get the sheet data of the worksheet
								SheetData sheetData = worksheetPart.Worksheet.GetFirstChild<SheetData>();

								// Add headers to the sheet data
								Row headerRow = new Row();
								headerRow.Append(
									CreateCell("Emploee Location"),
									CreateCell("Emploee Code"),
									CreateCell("Emploee Name"),
									CreateCell("Material Code"),
									CreateCell("Material Name"),
									CreateCell("Material Unit"),
									CreateCell("Kind"),
									CreateCell("Store"),
									CreateCell("Store Name"),
									CreateCell("Date"),
									CreateCell("Document Number"),
									CreateCell("Quantity"),
									CreateCell("Location")
								);
								sheetData.Append(headerRow);

								// Add data rows to the sheet data
								foreach (DataRow row in dt.Rows)
								{
									Row dataRow = new Row();
									dataRow.Append(
										CreateCell(row["EMP_LOC"].ToString()),
										CreateCell(row["EMP_CODE"].ToString()),
										CreateCell(row["EMP_NAME"].ToString()),
										CreateCell(row["MTRL_CODE"].ToString()),
										CreateCell(row["MTRL_NAME"].ToString()),
										CreateCell(row["MTRL_UNT"].ToString()),
										CreateCell(row["KIND"].ToString()),
										CreateCell(row["STOR"].ToString()),
										CreateCell(row["STOR_NAME"].ToString()),
										CreateCell(row["DTE"].ToString()),
										CreateCell(row["DOC_NO"].ToString()),
										CreateCell(row["QUT"].ToString()),
										CreateCell(row["LOC"].ToString())
									);
									sheetData.Append(dataRow);
								}

								// Add the worksheet to the workbook
								Sheets sheets = workbookPart.Workbook.AppendChild(new Sheets());
								Sheet sheet = new Sheet() { Id = workbookPart.GetIdOfPart(worksheetPart), SheetId = 1, Name = "Sheet1" };
								sheets.Append(sheet);

								// Save the workbook
								workbookPart.Workbook.Save();

								// Close the document
								document.Close();
							}
                            // Reset position before reading
                            stream.Position = 0;

                            // Return the Excel file to browser
                            return File(
                                stream.ToArray(),
                                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                "PersonalCustodyReport.xlsx"
                            );
							// Set the content type and file name for the response
							//Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
							//Response.Headers.Add("content-disposition", "attachment; filename=PersonalCustodyReport.xlsx");

							//// Write the Excel document to the response output stream
							//Response.Body.WriteAsync(stream.ToArray());
						}
						#endregion 2
						#region 1
						using (XLWorkbook wb = new XLWorkbook())
						{
							string sheetName = "Sheet1"; // Default sheet name
							if (!string.IsNullOrEmpty(dt.TableName))
							{
								sheetName = dt.TableName; // Use the DataTable name as the sheet name if available
							}
							wb.Worksheets.Add(dt, sheetName);
							using (MemoryStream stream = new MemoryStream())
							{
								wb.SaveAs(stream);
								try
								{
									string userProfileDirectory = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile);
									string filePath = Path.Combine(userProfileDirectory, "Downloads", "result.xlsx");
									File(stream.ToArray(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", filePath);
									//File(stream.ToArray(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "%USERPROFILE%\\Downloads\\result.xlsx");
								}
								catch (Exception ex)
								{
									Exception customEx = new Exception("error in downloading");
									throw customEx;
								}
							}
						}
						#endregion 1
						//return View("Done");
					}
				}
			}
		}
		public IActionResult PersonalCustodyCrystalReport()
		{
			return View(new AllMaterialsReportViewModel());
		}
		[Route("/Home/PersonalCustodyCrystalReportGetResult")]
		public IActionResult PersonalCustodyCrystalReportGetResult()
		{
			#region 1
			//ReportDocument rpt = new ReportDocument();
			//string path = @"E:\SSC Client\REPORT\ASSTP0060.rpt";
			//rpt.Load(path);
			//CrystalDecisions.CrystalReports.Engine.Database crDatabase = rpt.Database;
			//         CrystalDecisions.Shared.ConnectionInfo crConnectionInfo = new CrystalDecisions.Shared.ConnectionInfo();
			//crConnectionInfo.ServerName = "(local)";
			//crConnectionInfo.DatabaseName = "PLANDB";
			//crConnectionInfo.IntegratedSecurity = true;
			//rpt.ExportToDisk(ExportFormatType.Excel, "FilePath");
			#endregion 1
			#region 2
			//// Load the Crystal Report file
			//ReportClass rptH = new ReportClass();
			//List<sampledataset> data = objdb.getdataset();
			//rptH.FileName = Server.MapPath("[reportName].rpt");
			//rptH.Load();
			//rptH.SetDatabaseLogon("un", "pwd", "server", "db");
			//rptH.SetDataSource(data);
			//Stream stream = rptH.ExportToStream
			//   (CrystalDecisions.Shared.ExportFormatType.PortableDocFormat);
			//stream.Seek(0, System.IO.SeekOrigin.Begin);
			//return new FileStreamResult(stream, "application/pdf");
			#endregion 2
			#region 3
			// Create an instance of the Crystal Report document
			ReportDocument report = new ReportDocument();
			string reportPath = Path.Combine(_hostingEnvironment.WebRootPath, "Reports", "YourReport.rpt");
			report.Load(reportPath);

			// Set report parameters if needed
			// report.SetParameterValue("ParameterName", parameterValue);

			// Export the report to a PDF file
			Stream stream = report.ExportToStream(ExportFormatType.PortableDocFormat);

			// Clean up resources
			report.Close();
			report.Dispose();

			// Set the response content type and headers
			return File(stream, "application/pdf", "YourReport.pdf");
			#endregion 3
		}
        public ActionResult DownloadReport()
        {
            ReportDocument reportDocument = new ReportDocument();

            // Load the report document
            string basePath = AppDomain.CurrentDomain.BaseDirectory;
            string reportFileName = "YourReport.rpt";
            string rptFilePath = Path.Combine(basePath, "Reports", reportFileName);
            reportDocument.Load(rptFilePath);

            // Set database login credentials if needed
            // reportDocument.SetDatabaseLogon("username", "password", "server", "database");

            // Export the report to a file
            string exportFilePath = Path.Combine(basePath, "Reports", "YourReport.pdf");
            reportDocument.ExportToDisk(ExportFormatType.PortableDocFormat, exportFilePath);

            // Create a memory stream and read the exported file into it
            using (MemoryStream memoryStream = new MemoryStream())
            {
                using (FileStream fileStream = new FileStream(exportFilePath, FileMode.Open, FileAccess.Read))
                {
                    fileStream.CopyTo(memoryStream);
                }

                // Provide the memory stream for download
                memoryStream.Seek(0, SeekOrigin.Begin);
                return File(memoryStream, "application/pdf", "YourReport.pdf");
            }
        }
        public IActionResult Privacy()
		{
			return View();
		}

		[ResponseCache(Duration = 0, Location = ResponseCacheLocation.None, NoStore = true)]
		public IActionResult Error()
		{
			return View(new ErrorViewModel { RequestId = Activity.Current?.Id ?? HttpContext.TraceIdentifier });
		}
	}
}