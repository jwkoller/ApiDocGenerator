using Microsoft.Extensions.Logging;
using System.Text;

namespace APIDocGenerator.Services
{
    public class FileReaderService
    {
        private readonly ILogger<FileReaderService> _logger;

        public FileReaderService(ILogger<FileReaderService> logger) { 
            _logger = logger;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="directoryPath"></param>
        /// <returns></returns>
        public static IEnumerable<FileInfo> GetFiles(string directoryPath)
        {
            ArgumentException.ThrowIfNullOrWhiteSpace(directoryPath);
            IEnumerable<string> paths =  Directory.GetFiles(directoryPath, "*.cs", SearchOption.AllDirectories);
            return paths.Select(p => new FileInfo(p));
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="filePath"></param>
        /// <returns></returns>
        public static IEnumerable<string> GetValidFileLines(string filePath)
        {
            IEnumerable<string> lines = File.ReadAllLines(filePath).Select(x => x.Trim());

            return lines.Where(x => !string.IsNullOrWhiteSpace(x) && (x.StartsWith('[') || x.StartsWith("///")));
        }

        
    }
}
