// (C) Copyright 2019 by Autodesk, Inc. 
//
// Permission to use, copy, modify, and distribute this software
// in object code form for any purpose and without fee is hereby
// granted, provided that the above copyright notice appears in
// all copies and that both that copyright notice and the limited
// warranty and restricted rights notice below appear in all
// supporting documentation.
//
// AUTODESK PROVIDES THIS PROGRAM "AS IS" AND WITH ALL FAULTS. 
// AUTODESK SPECIFICALLY DISCLAIMS ANY IMPLIED WARRANTY OF
// MERCHANTABILITY OR FITNESS FOR A PARTICULAR USE.  AUTODESK,
// INC. DOES NOT WARRANT THAT THE OPERATION OF THE PROGRAM WILL
// BE UNINTERRUPTED OR ERROR FREE.
//
// Use, duplication, or disclosure by the U.S. Government is
// subject to restrictions set forth in FAR 52.227-19 (Commercial
// Computer Software - Restricted Rights) and DFAR 252.227-7013(c)
// (1)(ii)(Rights in Technical Data and Computer Software), as
// applicable.
//
// Revit FbxExporter
// by Eason Kang - Autodesk Forge & Autodesk Developer Network (ADN)
//

using System;
using System.IO;
using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using DesignAutomationFramework;
using System.Linq;
using System.Collections.Generic;
using Newtonsoft.Json;

namespace Autodesk.ADN.FbxExporter
{
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    public class MainApp : IExternalDBApplication
    {
        public ExternalDBApplicationResult OnStartup(ControlledApplication application)
        {
            DesignAutomationBridge.DesignAutomationReadyEvent += HandleDesignAutomationReadyEvent;
            return ExternalDBApplicationResult.Succeeded;
        }

        public ExternalDBApplicationResult OnShutdown(ControlledApplication application)
        {
            return ExternalDBApplicationResult.Succeeded;
        }

        public void HandleApplicationInitializedEvent(object sender, Autodesk.Revit.DB.Events.ApplicationInitializedEventArgs e)
        {
            var app = sender as Autodesk.Revit.ApplicationServices.Application;
            DesignAutomationData data = new DesignAutomationData(app, "InputFile.rvt");
            this.ExportFBX(data);
        }

        private void HandleDesignAutomationReadyEvent(object sender, DesignAutomationReadyEventArgs e)
        {
            LogTrace("Design Automation Ready event triggered...");
            e.Succeeded = true;
            e.Succeeded = this.ExportFBX(e.DesignAutomationData);
        }

        private bool ExportFBX(DesignAutomationData data)
        {
            if (data == null)
                return false;

            Application app = data.RevitApp;
            if (app == null)
            {
                LogTrace("Error occured");
                LogTrace("Invalid Revit App");
                return false;
            }

            string modelPath = data.FilePath;
            if (string.IsNullOrWhiteSpace(modelPath))
            {
                LogTrace("Error occured");
                LogTrace("Invalid File Path");
                return false;
            }

            var doc = data.RevitDoc;
            if (doc == null)
            {
                LogTrace("Error occured");
                LogTrace("Invalid Revit DB Document");
                return false;
            }

            var inputParams= JsonConvert.DeserializeObject<InputParams>(File.ReadAllText("params.json"));
            if (inputParams == null)
            {
                LogTrace("Invalid Input Params or Empty JSON Input");
                return false;
            }

            using (var collector = new FilteredElementCollector(doc))
            {
                LogTrace("Creating export folder...");

                var exportPath = Path.Combine(Directory.GetCurrentDirectory(), "exportedFBXs");
                if (!Directory.Exists(exportPath))
                {
                    try
                    {
                        Directory.CreateDirectory(exportPath);
                    }
                    catch (Exception ex)
                    {
                        this.PrintError(ex);
                        return false;
                    }
                }

                LogTrace(string.Format("Export Path: {0}", exportPath));

                LogTrace("Collecting 3D views...");
                IEnumerable<ElementId> viewIds = null;

                if (inputParams.ExportAll == true)
                {
                    viewIds = collector.WhereElementIsNotElementType()
                                        .OfClass(typeof(View3D))
                                        .WhereElementIsNotElementType()
                                        .Cast<View3D>()
                                        .Where(v => !v.IsTemplate)
                                        .Select(v => v.Id);
                }
                else
                {
                    try
                    {
                        if (inputParams.ViewIds == null || inputParams.ViewIds.Count() <= 0)
                        {
                            throw new InvalidDataException("Invalid input viewIds while the exportAll value is false!");
                        }

                        var viewElemIds = new List<ElementId>();
                        foreach (var viewGuid in inputParams.ViewIds)
                        {
                            var view = doc.GetElement(viewGuid);
                            if (view == null || (!(view is View3D)))
                                throw new InvalidDataException(string.Format("3D view not found with guid `{0}`", viewGuid));

                            viewElemIds.Add(view.Id);
                        }

                        viewIds = viewElemIds;
                    }
                    catch(Exception ex)
                    {
                        this.PrintError(ex);
                        return false;
                    }
                }

                if (viewIds == null || viewIds.Count() <= 0)
                {
                    LogTrace("Error occured");
                    LogTrace("No 3D views to be exported...");
                    return false;
                }

                LogTrace("Starting the export task...");

                try
                {
                    foreach (var viewId in viewIds)
                    {
                        var exportOpts = new FBXExportOptions();
                        exportOpts.StopOnError = true;
                        exportOpts.WithoutBoundaryEdges = true;

                        var view = doc.GetElement(viewId) as View3D;
                        var name = view.Name;
                        name = name.Replace("{", "_").Replace("}", "_").ReplaceInvalidFileNameChars();
                        var filename = string.Format("{0}.fbx", name);

                        LogTrace(string.Format("Exporting {0}...", filename));

                        var viewSet = new ViewSet();
                        viewSet.Insert(view);

                        doc.Export(exportPath, filename, viewSet, exportOpts);
                    }
                }
                catch (Autodesk.Revit.Exceptions.InvalidPathArgumentException ex)
                {
                    this.PrintError(ex);
                    return false;
                }
                catch (Autodesk.Revit.Exceptions.ArgumentException ex)
                {
                    this.PrintError(ex);
                    return false;
                }
                catch (Autodesk.Revit.Exceptions.InvalidOperationException ex)
                {
                    this.PrintError(ex);
                    return false;
                }
                catch (Exception ex)
                {
                    this.PrintError(ex);
                    return false;
                }
            }

            LogTrace("Exporting completed...");

            return true;
        }

        private void PrintError(Exception ex)
        {
            LogTrace("Error occured");
            LogTrace(ex.Message);

            if (ex.InnerException !=null)
                LogTrace(ex.InnerException.Message);
        }

        /// <summary>
        /// This will appear on the Design Automation output
        /// </summary>
        private static void LogTrace(string format, params object[] args)
        {
#if DEBUG
                System.Diagnostics.Trace.WriteLine(string.Format(format, args));
#endif
            System.Console.WriteLine(format, args);
        }
    }
}
