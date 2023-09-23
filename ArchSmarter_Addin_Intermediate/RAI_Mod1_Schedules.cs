#region Namespaces
using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Architecture;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Reflection;
using System.Windows;

#endregion

namespace ArchSmarter_Addin_Intermediate
{
    [Transaction(TransactionMode.Manual)]
    public class RAI_Mod1_Schedules : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            // this is a variable for the Revit application
            UIApplication uiapp = commandData.Application;

            // this is a variable for the current Revit model
            Document doc = uiapp.ActiveUIDocument.Document;

            //01. create a room schedule for each department

            FilteredElementCollector roomcollector = new FilteredElementCollector(doc);
            roomcollector.OfCategory(BuiltInCategory.OST_Rooms);
            

            using (Transaction T = new Transaction(doc))
            {
                T.Start("Create Room schedules by Department");
                {
                    ElementId catId = new ElementId(BuiltInCategory.OST_Rooms);
                    

                    //02. collect 1 instance of the Rooms to get the parameters
                    Element roomElement = roomcollector.FirstElement();

                    //03. get parameters
                    Parameter ParamRmlevel = roomElement.LookupParameter("Level");
                    Parameter ParamRmName = roomElement.get_Parameter(BuiltInParameter.ROOM_NAME);
                    Parameter ParamRmNum = roomElement.get_Parameter(BuiltInParameter.ROOM_NUMBER);
                    Parameter ParamRmDept = roomElement.get_Parameter(BuiltInParameter.ROOM_DEPARTMENT);
                    Parameter ParamRmArea = roomElement.get_Parameter(BuiltInParameter.ROOM_AREA);
                    //Parameter ParamRmComs = roomElement.LookupParameter("Comments");
                    Parameter ParamRmComs = roomElement.get_Parameter(BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS);
                    


                    //05. by department creation 
                    List<string> deptList = new List<string>();
                    foreach (Room rm in roomcollector)
                    {                       
                       string deptname = GetParameterValuebyName(rm, "Department");
                       deptList.Add(deptname);
                        
                    } 

                    //filter for unique items 
                    List<string> uniDeptList = deptList.Distinct().ToList();             

                    //06. create schedule
                    foreach (string dept in uniDeptList)
                    {
                        ViewSchedule DeptRoomSch = ViewSchedule.CreateSchedule(doc, catId);
                        DeptRoomSch.Name = "Dept - " + dept;

                        //04.create scheduelfields
                        ScheduleField FldRmLevel = DeptRoomSch.Definition.AddField(ScheduleFieldType.Instance, ParamRmlevel.Id);
                        ScheduleField FldRmName = DeptRoomSch.Definition.AddField(ScheduleFieldType.Instance, ParamRmName.Id);
                        ScheduleField FldRmNum = DeptRoomSch.Definition.AddField(ScheduleFieldType.Instance, ParamRmNum.Id);
                        ScheduleField FldRmDept = DeptRoomSch.Definition.AddField(ScheduleFieldType.Instance, ParamRmDept.Id);
                        ScheduleField FldRmArea = DeptRoomSch.Definition.AddField(ScheduleFieldType.ViewBased, ParamRmArea.Id);
                        ScheduleField FldRmComs = DeptRoomSch.Definition.AddField(ScheduleFieldType.Instance, ParamRmComs.Id);
                        ScheduleField FldRmCont = DeptRoomSch.Definition.AddField(ScheduleFieldType.Count);
                        FldRmLevel.IsHidden = true;
                        
                        //07. add schedule filter, then add it to the schedule via definition settings
                        ScheduleFilter schFilter = new ScheduleFilter(FldRmDept.FieldId, ScheduleFilterType.Equal, dept);
                        DeptRoomSch.Definition.AddFilter(schFilter);

                        //08A forgot to group by Level!
                        ScheduleSortGroupField sortLevel = new ScheduleSortGroupField(FldRmLevel.FieldId);
                        sortLevel.ShowHeader = true;
                        sortLevel.ShowFooter = true;
                        //make sure to shoe header or footer before adding definition!
                        DeptRoomSch.Definition.AddSortGroupField(sortLevel);
                        
                        //08. Sort by RoomName
                        ScheduleSortGroupField sortRoomName = new ScheduleSortGroupField(FldRmName.FieldId);
                        DeptRoomSch.Definition.AddSortGroupField(sortRoomName);
                        

                        //10. the toal area
                        
                        DeptRoomSch.Definition.ShowGrandTotal = true;

                        //09. add total area for level group
                        FldRmArea.DisplayType = ScheduleFieldDisplayType.Totals;
                        FldRmCont.DisplayType = ScheduleFieldDisplayType.Totals;

                    }

                    //11. create a schedule of departments 
                    ViewSchedule DeptSchAll = ViewSchedule.CreateSchedule(doc, catId);
                    //11a. set schedule to all depts
                    DeptSchAll.Name = "All Departments";
                    //11b. add only deptment name and areas.
                    ScheduleField deptName = DeptSchAll.Definition.AddField(ScheduleFieldType.Instance, ParamRmDept.Id);
                    ScheduleField deptArea = DeptSchAll.Definition.AddField(ScheduleFieldType.ViewBased, ParamRmArea.Id);
                    ScheduleField deptCont = DeptSchAll.Definition.AddField(ScheduleFieldType.Count);
                    //11c. display totals of areas
                    deptArea.DisplayType = ScheduleFieldDisplayType.Totals;
                    deptCont.DisplayType = ScheduleFieldDisplayType.Totals;
                    //11.d sort by department name 
                    ScheduleSortGroupField sortdeptName = new ScheduleSortGroupField(deptName.FieldId);
                    DeptSchAll.Definition.AddSortGroupField(sortdeptName);
                    //11.e non itemised
                    DeptSchAll.Definition.IsItemized = false;
                    //11.f show grand totals
                    DeptSchAll.Definition.ShowGrandTotal = true;
                    //11.g and title! why are these seperate.
                    DeptSchAll.Definition.ShowGrandTotalTitle = true;

                }

                T.Commit();
            }         

            return Result.Succeeded;
        }

        private string GetParameterValuebyName(Element element, string paramName)
        {            
            IList<Parameter> paramList = element.GetParameters(paramName);
            Parameter myParam = paramList.First();
            return myParam.AsString();
            
        }

        internal static PushButtonData GetButtonData()
        {
            // use this method to define the properties for this command in the Revit ribbon
            string buttonInternalName = "btnCommand1";
            string buttonTitle = "Button 1";

            ButtonDataClass myButtonData1 = new ButtonDataClass(
                buttonInternalName,
                buttonTitle,
                MethodBase.GetCurrentMethod().DeclaringType?.FullName,
                Properties.Resources.Blue_32,
                Properties.Resources.Blue_16,
                "This is a tooltip for Button 1");

            return myButtonData1.Data;
        }
    }
}
