// See https://aka.ms/new-console-template for more information
using Aspose.Email;
using Aspose.Email.Mapi;
using Aspose.Email.Storage.Pst;

Console.Write("Path: ");
string dirPath = Console.ReadLine();

var outFileName = $"{dirPath}\\outputFile.pst";

if (File.Exists(outFileName))
    File.Delete(outFileName);

using (var personalStorage = PersonalStorage.Create(outFileName, FileFormatVersion.Unicode))
{
    var directories = Directory.GetDirectories(dirPath);
    if (directories.Any())
    {
        foreach (var d in directories)
            SetFilesIntoBox(personalStorage, d);
    }
    else
        SetFilesIntoBox(personalStorage, dirPath);
}

void SetFilesIntoBox(PersonalStorage personalStorage, string directoryPath)
{
    var pathBox = personalStorage.RootFolder.AddSubFolder(Path.GetFileName(directoryPath));
    Console.WriteLine($"Create box: {pathBox.DisplayName}");
    foreach (var f in Directory.GetFiles(directoryPath, "*.eml"))
    {
        using (var message = MailMessage.Load(f))
        {
            pathBox.AddMessage(MapiMessage.FromMailMessage(message, MapiConversionOptions.UnicodeFormat));
            Console.WriteLine($"Add message: {Path.GetFileName(f)}");
        }
    }
}