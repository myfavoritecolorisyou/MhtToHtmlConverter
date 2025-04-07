using System;
using System.IO;
using System.Text.RegularExpressions;
using MimeKit;

class Program
{
    static void Main()
    {
        Console.WriteLine("프로그램 사용 전, ReadMe 정독 권장 드립니다.");
        Console.Write("MHT 파일이 있는 폴더 경로를 입력하세요: ");
        string inputFolder = Console.ReadLine() ?? @"C:\mht_files";

        Console.Write("변환된 HTML 파일을 저장할 폴더 경로를 입력하세요: ");
        string outputFolder = Console.ReadLine() ?? @"C:\html_output";

        // 입력 폴더 확인
        if (!Directory.Exists(inputFolder))
        {
            Console.WriteLine("입력 폴더를 찾을 수 없습니다.");
            return;
        }

        // 출력 폴더 생성
        if (!Directory.Exists(outputFolder))
        {
            Console.WriteLine($"{outputFolder} 폴더를 생성합니다.");
            Directory.CreateDirectory(outputFolder);
        }

        // 모든 파일 목록 가져오기 (하위 폴더 포함)
        string[] allFiles = Directory.GetFiles(inputFolder, "*", SearchOption.AllDirectories);
        if (allFiles.Length == 0)
        {
            Console.WriteLine("입력 폴더에 파일이 없습니다.");
            return;
        }

        Console.WriteLine($"총 {allFiles.Length}개의 파일을 처리합니다...");

        int successCount = 0;
        int copyCount = 0;

        foreach (var file in allFiles)
        {
            try
            {
                // 상대 경로 계산
                string relativePath = file[(inputFolder.Length + 1)..];
                string outputFilePath;

                if (Path.GetExtension(file).Equals(".mht", StringComparison.OrdinalIgnoreCase))
                {
                    // 변환된 HTML 파일 경로
                    string newFileName = Path.GetFileNameWithoutExtension(relativePath) + ".html";
                    string newDirPath = Path.Combine(outputFolder, Path.GetDirectoryName(relativePath) ?? "");
                    outputFilePath = Path.Combine(newDirPath, newFileName);

                    // 변환된 HTML 파일을 저장할 폴더 생성
                    if (!Directory.Exists(newDirPath))
                    {
                        Directory.CreateDirectory(newDirPath);
                    }

                    ConvertMhtToHtml(file, outputFilePath);
                    Console.WriteLine($"변환 완료: {outputFilePath}");
                    successCount++;
                }
                else
                {
                    // 다른 파일은 그대로 복사
                    outputFilePath = Path.Combine(outputFolder, relativePath);
                    string outputDir = Path.GetDirectoryName(outputFilePath);

                    // 복사할 폴더가 없으면 생성
                    if (!Directory.Exists(outputDir))
                    {
                        Directory.CreateDirectory(outputDir);
                    }
                    File.Copy(file, outputFilePath, true);
                    Console.WriteLine($"복사 완료: {outputFilePath}");
                    copyCount++;
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"오류 발생 ({file}): {ex.Message}");
            }
        }

        Console.WriteLine($"변환 완료. MHT 변환: {successCount}개, 일반 파일 복사: {copyCount}개.");
        Console.WriteLine("아무 키나 누르면 종료합니다...");
        Console.ReadKey();
    }

    private static void ConvertMhtToHtml(string mhtPath, string htmlPath)
    {
        string outputDir = Path.GetDirectoryName(htmlPath);
        if (!Directory.Exists(outputDir))
        {
            Directory.CreateDirectory(outputDir);
        }

        var mhtMessage = MimeMessage.Load(mhtPath);
        var htmlBody = mhtMessage.HtmlBody ?? throw new Exception("HTML 내용을 찾을 수 없습니다.");

        var cidToBase64Map = new Dictionary<string, string>();
        var locationToBase64Map = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

        foreach (var part in mhtMessage.BodyParts)
        {
            if (part is MimePart mimePart)
            {
                using var ms = new MemoryStream();
                mimePart.Content.DecodeTo(ms);
                string base64 = Convert.ToBase64String(ms.ToArray());
                string mime = mimePart.ContentType.MimeType;

                // CID 처리
                if (!string.IsNullOrEmpty(mimePart.ContentId))
                {
                    string cid = mimePart.ContentId.Trim('<', '>');
                    cidToBase64Map[cid] = $"data:{mime};base64,{base64}";
                }

                // Content-Location 처리
                if (mimePart.ContentLocation is not null)
                {
                    string location = mimePart.ContentLocation.ToString();
                    locationToBase64Map[location] = $"data:{mime};base64,{base64}";
                }
            }
        }
        
        htmlBody = cidToBase64Map.Aggregate(htmlBody, (current, kvp) => current.Replace($"cid:{kvp.Key}", kvp.Value));  // CID 이미지 대체
        htmlBody = locationToBase64Map.Aggregate(htmlBody, (current, kvp) => current.Replace(kvp.Key, kvp.Value));      // Content-Location 이미지 대체
        htmlBody = ReplaceFileSrcWithBase64(htmlBody); // file:/// 이미지 경로 처리

        File.WriteAllText(htmlPath, htmlBody);
    }


    // 파일 경로에서 base64 변환해 HTML에 반영
    private static string ReplaceFileSrcWithBase64(string html)
    {
        const string pattern = @"<img[^>]+src\s*=\s*[""']file:///([^""']+)[""'][^>]*>";
        return Regex.Replace(html, pattern, match =>
        {
            string fullPath = Uri.UnescapeDataString(match.Groups[1].Value);
            if (!File.Exists(fullPath)) return match.Value; // 파일이 없으면 그대로 둠

            try
            {
                // 파일을 읽어와서 base64로 인코딩
                byte[] imageBytes = File.ReadAllBytes(fullPath);
                string base64 = Convert.ToBase64String(imageBytes);
                
                // MIME 타입 결정
                string extension = Path.GetExtension(fullPath).ToLower();
                string mimeType = extension switch
                {
                    ".jpg" or ".jpeg" => "image/jpeg",
                    ".png" => "image/png",
                    ".gif" => "image/gif",
                    ".bmp" => "image/bmp",
                    _ => "application/octet-stream"
                };
                
                string dataUri = $"data:{mimeType};base64,{base64}";
                return Regex.Replace(match.Value, @"src\s*=\s*[""'][^""']+[""']", $"src=\"{dataUri}\"");
            }
            catch
            {
                return match.Value; // 오류 시 원본 유지
            }
        });
    }
}
