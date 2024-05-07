using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System.Collections.Generic;
using System.Linq;

namespace CutWallByDoor
{
    /// <summary>
    /// A command class for automating the process of cutting wall geometry at intersections with doors.
    /// </summary>
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    public class CommandCutWallByDoor : IExternalCommand
    {
        /// <summary>
        /// The main method of the command, invoked by Revit when the command is activated.
        /// </summary>
        /// <param name="commandData">Provides access to the Revit application and document objects.</param>
        /// <param name="message">A message that can be modified to return information back to Revit.</param>
        /// <param name="elements">A set of elements that can be modified to return information back to Revit.</param>
        /// <returns>The result of the command execution.</returns>
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIDocument uidoc = commandData.Application.ActiveUIDocument;
            Document doc = uidoc.Document;
            // Retrieve a list of all walls and doors in the document////
            List<Wall> walls = new FilteredElementCollector(doc).OfClass(typeof(Wall)).Cast<Wall>().ToList();
            List<FamilyInstance> doors = new FilteredElementCollector(doc).OfClass(typeof(FamilyInstance)).OfCategory(BuiltInCategory.OST_Doors).Cast<FamilyInstance>().ToList();

            using (Transaction trans = new Transaction(doc, "CutWallByDoors"))
            {
                trans.Start();

                foreach (FamilyInstance door in doors)
                {
                    BoundingBoxXYZ doorBox = door.get_BoundingBox(null);

                    foreach (Wall wall in walls)
                    {
                        BoundingBoxXYZ wallBox = wall.get_BoundingBox(null);

                        // Check the intersection of the bounding boxes of the wall and the door
                        if (BoundingBoxesIntersect(wallBox, doorBox))
                        {
                            try
                            {
                                // Perform the cut geometry operation
                                InstanceVoidCutUtils.AddInstanceVoidCut(doc, wall, door);
                            }
                            catch
                            {
                                // Exception handling
                            }
                        }
                    }
                }

                trans.Commit();
            }

            return Result.Succeeded;
        }

        /// <summary>
        /// Checks the intersection of two bounding boxes.
        /// </summary>
        /// <param name="box1">The first bounding box.</param>
        /// <param name="box2">The second bounding box.</param>
        /// <returns>Returns true if the boxes intersect.</returns>
        private bool BoundingBoxesIntersect(BoundingBoxXYZ box1, BoundingBoxXYZ box2)
        {
            bool separateX = box1.Max.X < box2.Min.X || box2.Max.X < box1.Min.X;
            bool separateY = box1.Max.Y < box2.Min.Y || box2.Max.Y < box1.Min.Y;
            bool separateZ = box1.Max.Z < box2.Min.Z || box2.Max.Z < box1.Min.Z;

            return !(separateX || separateY || separateZ);
        }
    }
}
