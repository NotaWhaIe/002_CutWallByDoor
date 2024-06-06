using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Timers;

using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Events;
using Autodesk.Revit.UI;

namespace CutWallByDoor
{
    [Autodesk.Revit.Attributes.TransactionAttribute(Autodesk.Revit.Attributes.TransactionMode.Manual)]
    public class DeleteLog
    {
        private Timer logTimer;
        private List<Tuple<string, string, int, string>> deletionLogBuffer = new List<Tuple<string, string, int, string>>();
        private readonly string sharedFolderPath = @"\\sb-sharegp\Bim2.0\5. Скрипты\999. BIM-отдел\RevitAutomation\DeleteLog";
        private readonly string usersFolder = @"\\sb-sharegp\Bim2.0\5. Скрипты\999. BIM-отдел\RevitAutomation\DeleteLog";
        private readonly object syncLock = new object();
        private Document currentDocument;
        private bool elementsDeleted = false;
        private string userPrefix = string.Empty;

        public void Initialize(UIControlledApplication application)
        {
            if (!IsUserAllowed().Result) return;

            application.ControlledApplication.DocumentOpened += OnDocumentOpened;
            application.ControlledApplication.DocumentChanged += HandleDocumentChanged;
            application.ControlledApplication.DocumentSynchronizedWithCentral += OnDocumentSynchronizedWithCentral;
            application.ControlledApplication.DocumentSaved += OnDocumentSaved;
            SetupTimer();
        }

        public void Shutdown(UIControlledApplication application)
        {
            logTimer.Stop();
            logTimer.Dispose();
            if (elementsDeleted)
            {
                SaveAndSync().Wait();
                MatchElements().Wait();
            }
        }

        private void OnDocumentOpened(object sender, DocumentOpenedEventArgs args)
        {
            currentDocument = args.Document;
            ExportProjectElements().Wait();
        }

        private void OnDocumentSynchronizedWithCentral(object sender, DocumentSynchronizedWithCentralEventArgs args)
        {
            currentDocument = args.Document;
            ExportProjectElements().Wait();
            if (elementsDeleted)
            {
                SaveAndSync().Wait();
                MatchElements().Wait();
            }
        }

        private void OnDocumentSaved(object sender, DocumentSavedEventArgs args)
        {
            currentDocument = args.Document;
            ExportProjectElements().Wait();
            if (elementsDeleted)
            {
                SaveAndSync().Wait();
                MatchElements().Wait();
            }
        }

        private async Task<bool> IsUserAllowed()
        {
            var userName = Environment.UserName.ToLowerInvariant();
            var computerName = Environment.MachineName.ToLowerInvariant();
            var userFiles = Directory.GetFiles(usersFolder, "*_users.txt");

            foreach (var userFile in userFiles)
            {
                var users = await Task.Run(() => File.ReadAllLines(userFile, Encoding.UTF8)
                            .Select(u => u.Trim().ToLowerInvariant())
                            .ToList());
                if (users.Contains(userName) || users.Contains(computerName))
                {
                    userPrefix = Path.GetFileNameWithoutExtension(userFile).Split('_')[0];
                    return true;
                }
            }
            return false;
        }

        private void HandleDocumentChanged(object sender, DocumentChangedEventArgs args)
        {
            Document doc = args.GetDocument();
            string projectName = Path.GetFileNameWithoutExtension(doc.PathName);
            string timeStamp = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
            string userName = doc.Application.Username;

            lock (syncLock)
            {
                foreach (ElementId id in args.GetDeletedElementIds())
                {
                    deletionLogBuffer.Add(new Tuple<string, string, int, string>(projectName, timeStamp, id.IntegerValue, userName));
                    elementsDeleted = true;
                }
            }
        }

        private void SetupTimer()
        {
            logTimer = new Timer(900000) { AutoReset = true, Enabled = true };
            logTimer.Elapsed += async (sender, e) =>
            {
                if (elementsDeleted)
                {
                    await SaveAndSync();
                    await MatchElements();
                }
            };
        }

        private async Task SaveAndSync()
        {
            if (currentDocument == null) return;

            string projectName = Path.GetFileNameWithoutExtension(currentDocument.PathName);
            string datePart = DateTime.Now.ToString("dd/MM/yyyy");
            string mainFolderName = $"{projectName}_{datePart.Replace("/", "-")}";
            string mainFolderPath = Path.Combine(sharedFolderPath, mainFolderName);
            Directory.CreateDirectory(mainFolderPath);

            string folderName = $"{projectName}_{userPrefix}";
            string folderPath = Path.Combine(mainFolderPath, folderName);
            string sourceTablesFolder = Path.Combine(folderPath, "SourceTables");
            Directory.CreateDirectory(folderPath);
            Directory.CreateDirectory(sourceTablesFolder);

            string filePath = Path.Combine(sourceTablesFolder, $"{folderName}.csv");

            List<Tuple<string, string, int, string>> bufferCopy;

            lock (syncLock)
            {
                bufferCopy = new List<Tuple<string, string, int, string>>(deletionLogBuffer);
                deletionLogBuffer.Clear();
            }

            await TrySaveToFile(async () =>
            {
                if (!File.Exists(filePath))
                {
                    await Task.Run(() => File.WriteAllText(filePath, "Project Name,Element ID,Time,User\n", Encoding.UTF8));
                }

                if (bufferCopy.Count > 0)
                {
                    await Task.Run(() => File.AppendAllLines(filePath, bufferCopy.Select(entry =>
                        $"{entry.Item1},{entry.Item3},{entry.Item2},{entry.Item4}"), Encoding.UTF8));
                }
            }, filePath);
        }

        private async Task MatchElements()
        {
            if (currentDocument == null) return;

            string projectName = Path.GetFileNameWithoutExtension(currentDocument.PathName);
            string datePart = DateTime.Now.ToString("dd/MM/yyyy");
            string mainFolderName = $"{projectName}_{datePart.Replace("/", "-")}";
            string mainFolderPath = Path.Combine(sharedFolderPath, mainFolderName);

            string folderName = $"{projectName}_{userPrefix}";
            string folderPath = Path.Combine(mainFolderPath, folderName);
            string sourceTablesFolder = Path.Combine(folderPath, "SourceTables");

            string filePathAR = Path.Combine(sourceTablesFolder, $"{folderName}.csv");
            string filePathDB = Path.Combine(sourceTablesFolder, $"{folderName}_Db.csv");
            string filePathARDeleted = Path.Combine(folderPath, $"{folderName}Deleted.csv");

            await TrySaveToFile(async () =>
            {
                if (!File.Exists(filePathAR) || !File.Exists(filePathDB)) return;

                var arData = await Task.Run(() => File.ReadAllLines(filePathAR, Encoding.UTF8))
                    .ContinueWith(task => task.Result.Skip(1)
                    .Select(line => line.Split(','))
                    .GroupBy(parts => int.Parse(parts[1]))
                    .ToDictionary(g => g.Key, g => g.First()));

                var dbData = await Task.Run(() => File.ReadAllLines(filePathDB, Encoding.UTF8))
                    .ContinueWith(task => task.Result.Skip(1)
                    .Select(line => line.Split(','))
                    .GroupBy(parts => int.Parse(parts[1]))
                    .ToDictionary(g => g.Key, g => g.First()));

                await TrySaveToFile(async () =>
                {
                    await Task.Run(() =>
                    {
                        using (var sw = new StreamWriter(filePathARDeleted, false, Encoding.UTF8))
                        {
                            sw.WriteLine("Project Name,Element ID,Element Type,Element Name,Level,Time,User");
                            foreach (var arEntry in arData)
                            {
                                if (dbData.TryGetValue(arEntry.Key, out var dbEntry))
                                {
                                    sw.WriteLine($"{arEntry.Value[0]},{arEntry.Key},{dbEntry[2]},{dbEntry[3]},{dbEntry[4]},{arEntry.Value[2]},{arEntry.Value[3]}");
                                }
                            }
                        }
                    });
                }, filePathARDeleted);

                elementsDeleted = false; // Сброс флага после обработки
            }, filePathARDeleted);
        }

        private async Task TrySaveToFile(Func<Task> saveAction, string filePath)
        {
            const int maxRetries = 2;
            int retries = 0;
            bool success = false;

            while (!success && retries < maxRetries)
            {
                try
                {
                    await saveAction();
                    success = true;
                }
                catch (IOException)
                {
                    retries++;
                    await Task.Delay(180000); // Ожидание 3 минуты перед повторной попыткой
                }
            }
        }

        private async Task ExportProjectElements()
        {
            if (currentDocument == null) return;

            string projectName = Path.GetFileNameWithoutExtension(currentDocument.PathName);
            string datePart = DateTime.Now.ToString("dd/MM/yyyy");
            string mainFolderName = $"{projectName}_{datePart.Replace("/", "-")}";
            string mainFolderPath = Path.Combine(sharedFolderPath, mainFolderName);
            Directory.CreateDirectory(mainFolderPath);

            string folderName = $"{projectName}_{userPrefix}";
            string folderPath = Path.Combine(mainFolderPath, folderName);
            string sourceTablesFolder = Path.Combine(folderPath, "SourceTables");
            Directory.CreateDirectory(folderPath);
            Directory.CreateDirectory(sourceTablesFolder);

            string filePath = Path.Combine(sourceTablesFolder, $"{folderName}_Db.csv");

            var collector = new FilteredElementCollector(currentDocument);
            collector.WhereElementIsNotElementType();

            var elementData = collector.Select(element =>
            {
                string elementType = element.Category?.Name ?? "Unknown";
                string elementName = element.Name;
                string level = element.LevelId != ElementId.InvalidElementId
                    ? currentDocument.GetElement(element.LevelId)?.Name
                    : "Unknown";

                return new Tuple<string, int, string, string, string>(
                    projectName, element.Id.IntegerValue, elementType, elementName, level);
            }).ToList();

            await TrySaveToFile(async () =>
            {
                await Task.Run(() =>
                {
                    using (var sw = new StreamWriter(filePath, false, Encoding.UTF8))
                    {
                        sw.WriteLine("Project Name,Element ID,Element Type,Element Name,Level");
                        foreach (var entry in elementData)
                        {
                            sw.WriteLine($"{entry.Item1},{entry.Item2},{entry.Item3},{entry.Item4},{entry.Item5}");
                        }
                    }
                });
            }, filePath);
        }
    }
}
