using Microsoft.AspNetCore.Mvc;
using System.IO;

[ApiController]
[Route("[controller]")]
public class SaveCallbackHandler : ControllerBase
{
    [HttpPost(Name = "FileUploader")]
    public async Task<IActionResult> ReceiveFile(IFormFile content)
    {
        if (content == null || content.Length <= 0)
        {
            return BadRequest("No file was uploaded.");
        }

        using (var memoryStream = new MemoryStream())
        {
            await content.CopyToAsync(memoryStream);
            byte[] fileBytes = memoryStream.ToArray();
            // Use the fileBytes as needed

            Console.WriteLine("File Read From Request Size - " + fileBytes.Length);
        }

        return Ok("File uploaded successfully.");
    }
}
