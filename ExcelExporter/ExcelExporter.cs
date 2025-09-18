using System.Text;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

namespace ExcelExporter
{
    public struct CellInfo
    {
        public string Type;
        public string Name;
        public string Desc;
    }

    public static class ExcelExporter
    {
        public static string[] appTypes = new string[7] { "Client", "Realm", "Gate", "Map", "Lobby", "DB", "Master"};
        public const string ExcelPath = "../Excel";
        private const string ServerConfigPath = "../Config/";

        public static bool isClient;
        public static void Main(string[] args)
        {
            try
            {
                string input;
                Console.WriteLine("輸出一般企劃表，選擇輸出端  [C]:Client   [S]:Server   [A]:All");
                input = Console.ReadLine().ToUpper();
                switch (input)
                {
                    case "C":
                        ExportClient();
                        break;
                    case "S":
                        ExportServer();
                        break;
                    case "A":
                        ExportClient();
                        ExportServer();
                        break;
                    default:
                        Console.WriteLine("別亂打好嗎，自己關掉重開，老子懶得幫你寫重選的方法");
                        break;
                }
                Console.ReadKey();
            }
            catch (Exception e)
            {
                throw new Exception(e.ToString());
            }

        }
        public static void ExportClient()
        {
            const string clientPath = "./Assets/Res/Config";

            isClient = true;

            Console.WriteLine("[1/3] 輸出企劃表txt檔...");
            ExportAll(clientPath);

            Console.WriteLine("[2/3] 輸出class檔(Model)...");
            ExportAllClass(@"./Assets/Model/Module/Demo/Config", "namespace ETModel\n{\n");

            Console.WriteLine("[3/3] 輸出class檔(Hotfix)...");
            ExportAllClass(@"./Assets/Hotfix/Module/Demo/Config", "using ETModel;\n\nnamespace ETHotfix\n{\n");

            Console.WriteLine($"輸出Client企劃表完成!");
        }
        public static void ExportServer()
        {
            isClient = false;

            Console.WriteLine("[1/2] 輸出企劃表txt檔...");
            ExportAll(ServerConfigPath);

            Console.WriteLine("[2/2] 輸出class檔(Server)...");
            ExportAllClass(@"../Server/Model/Module/Demo/Config", "namespace ETModel\n{\n");

            Console.WriteLine($"輸出Server企劃表完成!");
        }
        public static void ExportAllClass(string exportDir, string csHead)
        {
            string[] files = Directory.GetFiles(ExcelPath);
            foreach (string filePath in files)
            {
                if (Path.GetExtension(filePath) != ".xlsx")
                {
                    continue;
                }
                if (Path.GetFileName(filePath).StartsWith("~"))
                {
                    continue;
                }
                Console.WriteLine($"產生{Path.GetFileNameWithoutExtension(filePath)}類");
                ExportClass(filePath, exportDir, csHead);
            }
        }

        public static void ExportClass(string fileName, string exportDir, string csHead)
        {
            XSSFWorkbook xssfWorkbook;
            using (FileStream file = new FileStream(fileName, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            {
                xssfWorkbook = new XSSFWorkbook(file);
            }

            ISheet sheet = xssfWorkbook.GetSheetAt(0);

            String fieldAppType = GetCellString(sheet, 0, 0);
            // 本次輸出針對Client
            if (!fieldAppType.Contains(appTypes[0]) && isClient)
            {// 此Excel無Client標籤故不輸出
                return;
            }

            // 本次輸出針對Server
            if (!isClient)
            {
                for (int i = 1; i < appTypes.Length; i++)
                {
                    if (fieldAppType.Contains(appTypes[i]))
                    {
                        break;
                    }
                    else if (i == appTypes.Length - 1)
                    {// 此Excel無Server標籤故不輸出
                        return;
                    }
                }
            }
            string protoName = Path.GetFileNameWithoutExtension(fileName);

            string exportPath = Path.Combine(exportDir, $"{protoName}.cs");
            using (FileStream txt = new FileStream(exportPath, FileMode.Create))
            using (StreamWriter sw = new StreamWriter(txt))
            {
                StringBuilder sb = new StringBuilder();
                sb.Append(csHead);

                sb.Append($"\t[Config((int)({GetCellString(sheet, 0, 0)}))]\n");
                sb.Append($"\tpublic partial class {protoName}Category : ACategory<{protoName}>\n");
                sb.Append("\t{\n");
                sb.Append("\t}\n\n");

                sb.Append($"\tpublic class {protoName}: IConfig\n");
                sb.Append("\t{\n");
                sb.Append("\t\tpublic long Id { get; set; }\n");

                int cellCount = sheet.GetRow(3).LastCellNum;
                var sbLog = new StringBuilder();
                sbLog.Append($"{protoName}");
                for (int i = 2; i < cellCount; i++)
                {
                    string fieldDesc = GetCellString(sheet, 2, i);

                    if (fieldDesc.StartsWith("#"))
                    {
                        continue;
                    }

                    // s開頭代表這個是Server專用
                    if (fieldDesc.StartsWith("s") && isClient)
                    {
                        continue;
                    }

                    string fieldName = GetCellString(sheet, 3, i);

                    if (string.IsNullOrEmpty(fieldName))
                    {
                        continue;
                    }

                    if (fieldName == "Id" || fieldName == "_id")
                    {
                        continue;
                    }

                    string fieldType = GetCellString(sheet, 4, i);
                    if (fieldType == "" || fieldName == "")
                    {
                        continue;
                    }

                    sb.Append($"\t\tpublic {fieldType} {fieldName};\n");
                }
                if (sbLog.Length > protoName.Length)
                {
                    Console.ForegroundColor = ConsoleColor.Red;
                    Console.WriteLine(sbLog.ToString());
                    Console.ResetColor();
                    Console.ReadKey();
                    return;
                }
                sb.Append("\t}\n");
                sb.Append("}\n");

                sw.Write(sb.ToString());
            }
        }


        public static void ExportAll(string exportDir)
        {
            string[] files = Directory.GetFiles(ExcelPath);

            foreach (string filePath in files)
            {
                if (Path.GetExtension(filePath) != ".xlsx")
                {
                    continue;
                }
                if (Path.GetFileName(filePath).StartsWith("~"))
                {
                    continue;
                }
                Export(filePath, exportDir);
            }

            Console.WriteLine("所有企劃表讀取完成!");
        }

        public static void Export(string fileName, string exportDir)
        {
            XSSFWorkbook xssfWorkbook;
            string fieldAppType;
            using (FileStream file = new FileStream(fileName, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            {
                xssfWorkbook = new XSSFWorkbook(file);
            }
            string protoName = Path.GetFileNameWithoutExtension(fileName);
            string exportPath = Path.Combine(exportDir, $"{protoName}.txt");

            ISheet sheet = xssfWorkbook.GetSheetAt(0);
            fieldAppType = GetCellString(sheet, 0, 0);
            // 本次輸出針對Client
            if (!fieldAppType.Contains(appTypes[0]) && isClient)
            {// 此Excel無Client標籤故不輸出
                return;
            }

            // 本次輸出針對Server
            if (!isClient)
            {
                for (int i = 1; i < appTypes.Length; i++)
                {
                    if (fieldAppType.Contains(appTypes[i]))
                    {
                        break;
                    }
                    else if (i == appTypes.Length - 1)
                    {// 此Excel無Server標籤故不輸出
                        return;
                    }
                }
            }

            using (FileStream txt = new FileStream(exportPath, FileMode.Create))
            using (StreamWriter sw = new StreamWriter(txt))
            {
                Console.WriteLine($"{protoName}.txt 輸出開始");
                ExportSheet(sheet, sw);
            }

            Console.WriteLine($"{protoName}輸出完成");
        }

        public static void ExportSheet(ISheet sheet, StreamWriter sw)
        {
            int cellCount = sheet.GetRow(3).LastCellNum;
            CellInfo[] cellInfos = new CellInfo[cellCount];
            for (int i = 2; i < cellCount; i++)
            {
                string fieldDesc = GetCellString(sheet, 2, i);
                string fieldName = GetCellString(sheet, 3, i);
                string fieldType = GetCellString(sheet, 4, i);
                cellInfos[i] = new CellInfo() { Name = fieldName, Type = fieldType, Desc = fieldDesc };
            }

            for (int i = 5; i <= sheet.LastRowNum; ++i)
            {
                if (GetCellString(sheet, i, 2) == "")
                {
                    continue;
                }
                StringBuilder sb = new StringBuilder();
                sb.Append("{");
                IRow row = sheet.GetRow(i);
                for (int j = 2; j < cellCount; ++j)
                {
                    string desc = cellInfos[j].Desc.ToLower();

                    if (string.IsNullOrEmpty(desc))
                    {
                        continue;
                    }

                    if (desc.StartsWith("#"))
                    {
                        continue;
                    }

                    string fieldValue = GetCellString(row, j);
                    if (fieldValue == "")
                    {
                        Console.ForegroundColor = ConsoleColor.Red;
                        Console.WriteLine($"sheet: {sheet.SheetName} 中有空白數值 (第{i + 1}橫排, 第{j + 1}格, 欄位 {desc})");
                        Console.Read();
                        throw new Exception();
                    }

                    if (j > 2)
                    {
                        sb.Append(",");
                    }

                    string fieldName = cellInfos[j].Name;

                    if (fieldName == "Id" || fieldName == "_id")
                    {
                        if (isClient)
                        {
                            fieldName = "Id";
                        }
                        else
                        {
                            fieldName = "_id";
                        }
                    }

                    string fieldType = cellInfos[j].Type;
                    sb.Append($"\"{fieldName}\":{Convert(fieldType, fieldValue)}");
                }
                sb.Append("}");
                sw.WriteLine(sb.ToString());
            }
        }
        public static string Convert(string type, string value)
        {
            switch (type)
            {
                case "int[]":
                case "int32[]":
                case "long[]":
                //case "float[]": bson不支援float
                case "double[]":
                    return $"[{value}]";
                case "string[]":
                    return $"[{value}]";
                case "int":
                case "int32":
                case "int64":
                case "long":
                //case "float": bson不支援float
                case "double":
                    return value;
                case "string":
                    return $"\"{value}\"";
                default:
                    Console.ForegroundColor = ConsoleColor.Red;
                    Console.WriteLine($"企劃表不支援此類型: {type}");
                    Console.ReadKey();
                    throw new Exception();
            }
        }

        public static string GetCellString(ISheet sheet, int i, int j)
        {
            return sheet.GetRow(i)?.GetCell(j)?.ToString() ?? "";
        }

        public static string GetCellString(IRow row, int i)
        {
            return row?.GetCell(i)?.ToString() ?? "";
        }
    }
}

