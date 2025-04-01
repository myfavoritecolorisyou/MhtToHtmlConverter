using System;
using System.IO;
using MimeKit;

class Program
{
    static void Main()
    {
        // string inputFolder = @"D:\mht_files";   // 변환할 MHT 파일이 있는 폴더
        // string outputFolder = @"D:\html_output"; // 변환된 HTML 파일을 저장할 폴더
        Console.WriteLine("프로그램 사용 전, ReadMe 정독 권장 드립니다.");
        Console.Write("MHT 파일이 있는 폴더 경로를 입력하세요: ");
        string inputFolder = Console.ReadLine() ?? @"C:\mht_files";
        
        Console.Write("변환된 HTML 파일을 저장할 폴더 경로를 입력하세요: ");
        string outputFolder = Console.ReadLine() ?? @"C:\html_output"; ;

        // 입력 폴더가 없으면 종료
        if (!Directory.Exists(inputFolder))
        {
            Console.WriteLine("입력 폴더를 찾을 수 없습니다.");
            return;
        }

        // 출력 폴더가 없으면 생성
        if (!Directory.Exists(outputFolder))
        {
            Console.WriteLine($"{outputFolder} 폴더를 생성합니다.");
            Directory.CreateDirectory(outputFolder);
        }

        // MHT 파일이 포함된 모든 서브폴더 검색
        string[] mhtFiles = Directory.GetFiles(inputFolder, "*.mht", SearchOption.AllDirectories);
        if (mhtFiles.Length == 0)
        {
            Console.WriteLine("MHT 파일을 찾을 수 없습니다.");
            return;
        }

        Console.WriteLine($"총 {mhtFiles.Length}개의 MHT 파일을 변환합니다...");

        int successCount = 0;
        foreach (var mhtFile in mhtFiles)
        {
            try
            {
                // MHT 파일이 위치한 폴더 경로 계산
                string relativePath = mhtFile.Substring(inputFolder.Length + 1);  // 상대 경로 계산
                string outputFilePath = Path.Combine(outputFolder, Path.ChangeExtension(relativePath, ".html"));

                // HTML 파일을 저장할 디렉토리가 없으면 생성
                string outputDir = Path.GetDirectoryName(outputFilePath);
                if (!Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // MHT 파일을 HTML로 변환하여 해당 경로에 저장
                ConvertMhtToHtml(mhtFile, outputFilePath);
                Console.WriteLine($"변환 완료: {outputFilePath}");
                successCount++;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"오류 발생 ({mhtFile}): {ex.Message}");
            }
        }

        Console.WriteLine($"변환 완료. {successCount}/{mhtFiles.Length}개 파일 성공.");
    }

    static void ConvertMhtToHtml(string mhtPath, string htmlPath)
    {
        var mhtMessage = MimeMessage.Load(mhtPath);
        var htmlBody = mhtMessage.HtmlBody;

        if (string.IsNullOrEmpty(htmlBody))
        {
            throw new Exception("HTML 내용을 찾을 수 없습니다.");
        }

        // HTML 파일로 저장
        File.WriteAllText(htmlPath, htmlBody);
    }
}
