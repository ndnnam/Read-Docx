namespace OpenXMLSample.Helper
{
    public class FileHelper
    {
        public static async Task<string> UpLoad(IFormFile file)
        {
            var fileName = $"{Guid.NewGuid()}__{file.FileName}.docx";
            //var path = Path.Combine(Directory.GetCurrentDirectory(), "uploads");
            string fullPath = "uploads\\" + fileName;
            using(var stream = File.Create(fullPath))
            {
                await file.CopyToAsync(stream);
            }
            return fileName;
        }
    }
}
