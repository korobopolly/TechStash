using System;
using System.IO;
using OfficeOpenXml;
using System.Linq;
using System.Collections.Generic;
using Microsoft.VisualBasic.FileIO;  // 휴지통

namespace TechStash
{
    class Program
    {
        static void Main(string[] args)
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            string userProfile = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile);
            string downloadsPath = Path.Combine(userProfile, "Downloads");
            string folderPath = downloadsPath;

            var excelFiles = Directory.GetFiles(folderPath, "*.xlsx")
                                      .Where(f => Path.GetFileName(f).StartsWith("merge_"))
                                      .ToArray();

            if (excelFiles.Length == 0)
            {
                Console.WriteLine("조건을 만족하는 엑셀 파일이 없습니다.");
                return;
            }

            string mergedFileName = "Merged_" + DateTime.Now.ToString("yyyyMMddHHmmss");
            if (excelFiles.Length > 0)
            {
                string firstFileName = Path.GetFileNameWithoutExtension(excelFiles[0]);
                string[] parts = firstFileName.Split('_');
                if (parts.Length >= 3)
                {
                    mergedFileName = string.Join("_", parts, 1, parts.Length - 3) + "_" + DateTime.Now.ToString("yyyyMMddHHmmss");
                }
            }

            string mergedFilePath = Path.Combine(downloadsPath, mergedFileName + ".xlsx");

            // 시트순서 설정
            List<string> desiredOrder = new List<string> { "액세스룰", "인터페이스", "존", "네트워크객체", "네트워크그룹객체", "서비스객체", "서비스그룹객체", "일정객체", "NAT", "VPN" };

            Dictionary<ExcelWorksheet, int> sheetOrdering = new Dictionary<ExcelWorksheet, int>();

            using (ExcelPackage mergedPackage = new ExcelPackage())
            {
                foreach (var file in excelFiles)
                {
                    string fileName = Path.GetFileNameWithoutExtension(file);
                    string[] fileParts = fileName.Split('_');

                    string baseSheetName = fileParts.Length >= 2 ? fileParts[fileParts.Length - 2] : fileName;
                    string orderString = fileParts.Length >= 1 ? fileParts[fileParts.Length - 1] : "";

                    // 시트순서를 찾고, 없으면 큰 값 부여 (기본값 999)
                    int orderIndex = desiredOrder.IndexOf(baseSheetName);
                    orderIndex = (orderIndex == -1) ? 999 : orderIndex;

                    using (ExcelPackage package = new ExcelPackage(new FileInfo(file)))
                    {
                        int sheetCount = package.Workbook.Worksheets.Count;
                        for (int i = 0; i < sheetCount; i++)
                        {
                            var sheet = package.Workbook.Worksheets[i];
                            string newSheetName = (sheetCount == 1) ? baseSheetName : $"{baseSheetName}_{sheet.Name}";

                            if (mergedPackage.Workbook.Worksheets[newSheetName] != null)
                            {
                                int duplicateIndex = 1;
                                while (mergedPackage.Workbook.Worksheets[$"{newSheetName}_{duplicateIndex}"] != null)
                                {
                                    duplicateIndex++;
                                }
                                newSheetName = $"{newSheetName}_{duplicateIndex}";
                            }

                            var newSheet = mergedPackage.Workbook.Worksheets.Add(newSheetName);
                            if (sheet.Dimension != null)
                            {
                                var sourceRange = sheet.Cells[sheet.Dimension.Address];
                                var destinationCell = newSheet.Cells[1, 1];
                                sourceRange.Copy(destinationCell);
                            }

                            sheetOrdering[newSheet] = orderIndex;
                        }
                    }
                }

                using (ExcelPackage reorderedPackage = new ExcelPackage())
                {
                    var orderedWorksheets = mergedPackage.Workbook.Worksheets
                        .Cast<ExcelWorksheet>()
                        .OrderBy(ws => sheetOrdering[ws])
                        .ToList();

                    foreach (var sheet in orderedWorksheets)
                    {
                        var newSheet = reorderedPackage.Workbook.Worksheets.Add(sheet.Name);
                        if (sheet.Dimension != null)
                        {
                            var sourceRange = sheet.Cells[sheet.Dimension.Address];
                            var destinationCell = newSheet.Cells[1, 1];
                            sourceRange.Copy(destinationCell);
                        }
                    }

                    FileInfo fi = new FileInfo(mergedFilePath);
                    reorderedPackage.SaveAs(fi);
                }
            }

            Console.WriteLine("엑셀 파일 병합이 완료되었습니다. 저장 경로: " + mergedFilePath);

            // 원본 merge_ 엑셀 파일 (휴지통)
            MoveFilesToRecycleBin(excelFiles);
        }
        static void DeleteOriginalExcelFiles(string[] files)
        {
            foreach (var file in files)
            {
                try
                {
                    File.Delete(file);
                    Console.WriteLine($"삭제 완료: {file}");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"삭제 실패: {file} - 오류: {ex.Message}");
                }
            }
        }

        static void MoveFilesToRecycleBin(string[] files)
        {
            foreach (var file in files)
            {
                try
                {
                    FileSystem.DeleteFile(file, UIOption.OnlyErrorDialogs, RecycleOption.SendToRecycleBin);
                    Console.WriteLine($"휴지통으로 이동 완료: {file}");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"휴지통 이동 실패: {file} - 오류: {ex.Message}");
                }
            }
        }
    }
}
