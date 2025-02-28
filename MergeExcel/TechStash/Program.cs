using System;
using System.IO;
using OfficeOpenXml;
using System.Linq;
using System.Collections.Generic;

namespace TechStash
{
    class Program
    {
        static void Main(string[] args)
        {
            // EPPlus 라이선스 컨텍스트 설정 (비상업적 사용)
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            // 사용자 홈 폴더의 Downloads 폴더 경로 가져오기
            string userProfile = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile);
            string downloadsPath = Path.Combine(userProfile, "Downloads");
            // Excel 파일들이 바로 Downloads 폴더에 있다고 가정
            string folderPath = downloadsPath;

            // 폴더 내의 모든 .xlsx 파일 가져오기
            var excelFiles = Directory.GetFiles(folderPath, "*.xlsx");

            // 병합 결과 파일 이름 결정: 첫 번째 파일의 이름을 '_' 기준으로 분리한 후,
            // 처음부터 마지막에서 3번째 요소까지를 결합 (즉, parts[0] ~ parts[parts.Length - 3])
            string mergedFileName = "Merged";
            if (excelFiles.Length > 0)
            {
                string firstFileName = Path.GetFileNameWithoutExtension(excelFiles[0]);
                string[] parts = firstFileName.Split('_');
                if (parts.Length >= 3)
                {
                    mergedFileName = string.Join("_", parts, 0, parts.Length - 2);
                }
            }
            // 병합 결과 파일 경로 (Downloads 폴더에 저장)
            string mergedFilePath = Path.Combine(downloadsPath, mergedFileName + ".xlsx");

            // 시트별 정렬을 위한 딕셔너리 (키: 시트, 값: 파일명에서 "_"로 분리한 마지막 요소)
            Dictionary<ExcelWorksheet, string> sheetOrdering = new Dictionary<ExcelWorksheet, string>();

            // 임시로 병합 결과를 저장할 패키지 생성
            using (ExcelPackage mergedPackage = new ExcelPackage())
            {
                foreach (var file in excelFiles)
                {
                    // 각 파일의 이름(확장자 제외)
                    string fileName = Path.GetFileNameWithoutExtension(file);
                    string[] fileParts = fileName.Split('_');

                    // 시트 이름: 파일명에서 마지막에서 2번째 요소 사용 (예: "액세스룰")
                    string baseSheetName = fileParts.Length >= 2 ? fileParts[fileParts.Length - 2] : fileName;
                    // 정렬 기준: 파일명에서 "_" 분리 시 가장 마지막 요소 (예: "20250228163241")
                    string orderString = fileParts.Length >= 1 ? fileParts[fileParts.Length - 1] : "";

                    using (ExcelPackage package = new ExcelPackage(new FileInfo(file)))
                    {
                        int sheetCount = package.Workbook.Worksheets.Count;
                        for (int i = 0; i < sheetCount; i++)
                        {
                            var sheet = package.Workbook.Worksheets[i];
                            // 파일에 시트가 한 개이면 baseSheetName 그대로, 여러 개이면 baseSheetName과 원래 시트명을 결합
                            string newSheetName = (sheetCount == 1) ? baseSheetName : $"{baseSheetName}_{sheet.Name}";

                            // 동일한 시트 이름이 이미 있으면 고유하게 변경
                            if (mergedPackage.Workbook.Worksheets[newSheetName] != null)
                            {
                                int duplicateIndex = 1;
                                while (mergedPackage.Workbook.Worksheets[$"{newSheetName}_{duplicateIndex}"] != null)
                                {
                                    duplicateIndex++;
                                }
                                newSheetName = $"{newSheetName}_{duplicateIndex}";
                            }

                            // 새로운 시트 생성 후 원본 시트 데이터 복사
                            var newSheet = mergedPackage.Workbook.Worksheets.Add(newSheetName);
                            if (sheet.Dimension != null)
                            {
                                var sourceRange = sheet.Cells[sheet.Dimension.Address];
                                var destinationCell = newSheet.Cells[1, 1];
                                sourceRange.Copy(destinationCell);
                            }

                            // 해당 시트의 정렬 기준 값 저장
                            sheetOrdering[newSheet] = orderString;
                        }
                    }
                }

                // 병합된 패키지의 시트를 정렬하기 위해 새 패키지를 생성
                using (ExcelPackage reorderedPackage = new ExcelPackage())
                {
                    // 파일명 마지막 요소(orderString)를 기준으로 오름차순 정렬
                    var orderedWorksheets = mergedPackage.Workbook.Worksheets
                        .Cast<ExcelWorksheet>()
                        .OrderBy(ws => sheetOrdering[ws])
                        .ToList();

                    // 정렬된 순서대로 시트를 새 패키지에 복사
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

                    // 재정렬된 패키지를 저장
                    FileInfo fi = new FileInfo(mergedFilePath);
                    reorderedPackage.SaveAs(fi);
                }
            }

            Console.WriteLine("엑셀 파일 병합이 완료되었습니다. 저장 경로: " + mergedFilePath);
        }
    }
}