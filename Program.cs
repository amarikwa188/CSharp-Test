string name, id, dirname;

Console.WriteLine("===== Office File Organizer =====");

Console.Write("Enter your name:\n>");
name = Console.ReadLine();
Console.Write("Enter  your student ID:\n>");
id  = Console.ReadLine();
Console.Write("Enter the directory path to organize:\n>");
dirname = Console.ReadLine();

string curpath = Directory.GetCurrentDirectory() + "\\" + dirname;
List<string> files = new List<string>();

try{
    files = new List<string>(Directory.EnumerateFiles(curpath));
}catch{
    Console.WriteLine($"'{dirname}' directory does not exist.");
}

string worddir = curpath + "\\WordFiles";
string exceldir = curpath + "\\ExcelFiles";
string pptdir = curpath + "\\PPTFiles";

Directory.CreateDirectory(worddir);
Directory.CreateDirectory(exceldir);
Directory.CreateDirectory(pptdir);

Console.WriteLine("\nOrganizing files...\n");

string extension;
FileInfo info;

const string resulttxt = "SummaryReport.txt";
string summary;

using (StreamWriter sw = File.CreateText(resulttxt)){
    sw.WriteLine($"Student: {name} (ID: {id})");
    sw.WriteLine($"Organized Files:");

    foreach (string file in files){
        extension = Path.GetExtension(file).ToLower();
        info = new FileInfo(file);
        switch(extension){
            case ".docx":
                Console.WriteLine($"Moved: {info.Name} -> WordFiles");
                sw.WriteLine($"- {info.Name} (Word, {info.Length/1000} KB, Created: {File.GetCreationTime(file)})");
                File.Move(curpath + "\\" + info.Name, worddir + $"\\{info.Name}");
                break;
            case ".xlsx":
                Console.WriteLine($"Moved: {info.Name} -> ExcelFiles");
                sw.WriteLine($"- {info.Name} (Excel, {info.Length/1000} KB, Created: {File.GetCreationTime(file)})");
                File.Move(curpath + "\\" + info.Name, exceldir + $"\\{info.Name}");
                break;
            case ".pptx":
                Console.WriteLine($"Moved: {info.Name} -> PPTFiles");
                sw.WriteLine($"- {info.Name} (PowerPoint, {info.Length/1000} KB, Created: {File.GetCreationTime(file)})");
                File.Move(curpath + "\\" + info.Name, pptdir + $"\\{info.Name}");
                break;
            default:
                break;
        }
    }

    Console.WriteLine($"\nSummary report created: {resulttxt}");
    Console.WriteLine("-----------------------");
}

summary = File.ReadAllText(resulttxt);
Console.WriteLine(summary);