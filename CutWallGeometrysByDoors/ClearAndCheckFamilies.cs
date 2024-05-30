using System;
using System.Collections.Generic;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System.Linq;

namespace CutWallByDoor
{
    [Autodesk.Revit.Attributes.TransactionAttribute(Autodesk.Revit.Attributes.TransactionMode.Manual)]
    public class ClearAndCheckFamilies : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            // Получаем текущий документ
            UIDocument uiDoc = commandData.Application.ActiveUIDocument;
            Document doc = uiDoc.Document;


            // Определяем GUID параметров
            Guid widthGuid = new Guid("8f2e4f93-9472-4941-a65d-0ac468fd6a5d");
            Guid heightGuid = new Guid("da753fe3-ecfa-465b-9a2c-02f55d0c2ff1");
            Guid levelMarkGuid = new Guid("6ec2f9e9-3d50-4d75-a453-26ef4e6d1625");
            Guid functionGuid = new Guid("d1f527c4-8806-4e6f-8368-bc581b3e730d");
            BuiltInParameter markParam = BuiltInParameter.ALL_MODEL_MARK;

            // Получаем все экземпляры семейств категории "Обобщенная модель" с именем "SD_Отверстие в стене_Прямоугольное"
            var familyInstances = new FilteredElementCollector(doc)
                .OfCategory(BuiltInCategory.OST_GenericModel)
                .OfClass(typeof(FamilyInstance))
                .Cast<FamilyInstance>()
                .Where(fi => fi.Symbol.FamilyName.Equals("SD_Отверстие в стене_Прямоугольное"))
                .ToList();

            // Создаем словарь для хранения коллекций по марке
            var markCollections = new Dictionary<string, List<FamilyInstance>>();

            using (Transaction trans = new Transaction(doc, "Clear and Check Families"))
            {
                trans.Start();

                // Очищаем параметр "Примечание" у всех экземпляров
                foreach (var instance in familyInstances)
                {
                    var noteParam = instance.LookupParameter("Примечание");
                    if (noteParam != null)
                    {
                        noteParam.Set(string.Empty);
                    }

                    // Собираем экземпляры семейств в коллекции по параметру "Марка"
                    var markParamValue = instance.get_Parameter(markParam);
                    if (markParamValue != null && !string.IsNullOrEmpty(markParamValue.AsString()))
                    {
                        var mark = markParamValue.AsString();
                        if (!markCollections.ContainsKey(mark))
                        {
                            markCollections[mark] = new List<FamilyInstance>();
                        }
                        markCollections[mark].Add(instance);
                    }
                }

                // Проверяем параметры и устанавливаем примечание "Марка дублируется"
                foreach (var collection in markCollections.Values)
                {
                    if (collection.Count > 1)
                    {
                        bool markDuplicated = false;
                        var referenceInstance = collection[0];

                        foreach (var instance in collection.Skip(1))
                        {
                            if (!AreParametersEqual(referenceInstance, instance, widthGuid, heightGuid, levelMarkGuid, functionGuid))
                            {
                                markDuplicated = true;
                                break;
                            }
                        }

                        if (markDuplicated)
                        {
                            foreach (var instance in collection)
                            {
                                var noteParam = instance.LookupParameter("Примечание");
                                if (noteParam != null)
                                {
                                    noteParam.Set("Марка дублируется");
                                }
                            }
                        }
                    }
                }

                trans.Commit();
            }

            return Result.Succeeded;
        }

        private bool AreParametersEqual(FamilyInstance firstInstance, FamilyInstance secondInstance, Guid widthGuid, Guid heightGuid, Guid levelMarkGuid, Guid functionGuid)
        {
            var paramsToCheck = new List<Guid>
            {
                widthGuid,
                heightGuid,
                levelMarkGuid,
                functionGuid
            };
            
            foreach (var paramGuid in paramsToCheck)
            {
                var firstParam = firstInstance.get_Parameter(paramGuid);
                var secondParam = secondInstance.get_Parameter(paramGuid);

                if (firstParam == null || secondParam == null)
                {
                    return false;
                }

                if (firstParam.StorageType == StorageType.Double && secondParam.StorageType == StorageType.Double)
                {
                    if (firstParam.AsDouble() != secondParam.AsDouble())
                    {
                        return false;
                    }
                }
                else if (firstParam.StorageType == StorageType.String && secondParam.StorageType == StorageType.String)
                {
                    if (firstParam.AsString() != secondParam.AsString())
                    {
                        return false;
                    }
                }
                else if (firstParam.StorageType == StorageType.Integer && secondParam.StorageType == StorageType.Integer)
                {
                    if (firstParam.AsInteger() != secondParam.AsInteger())
                    {
                        return false;
                    }
                }
            }

            return true;
        }
    }
}
