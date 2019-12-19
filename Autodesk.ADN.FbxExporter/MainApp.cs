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
                return false;

            string modelPath = data.FilePath;
            if (string.IsNullOrWhiteSpace(modelPath))
                return false;

            var doc = data.RevitDoc;
            if (doc == null)
                return false;

            using (var collector = new FilteredElementCollector(doc))
            {
                LogTrace("Collecting 3D views...");

                var exportPath = Path.Combine(Directory.GetCurrentDirectory(), "exported");
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

                var veiwIds = collector.WhereElementIsNotElementType()
                                        .OfClass(typeof(View3D))
                                        .WhereElementIsNotElementType()
                                        .Cast<View3D>()
                                        .Where(v => !v.IsTemplate)
                                        .Select(v => v.Id);

                if (veiwIds == null || veiwIds.Count() <= 0)
                {
                    LogTrace("No 3D views to be exported...");
                    return false;
                }

                LogTrace("Starting the export task...");

                try
                {
                    foreach (var viewId in veiwIds)
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

            return true;
        }

        private void PrintError(Exception ex)
        {
            LogTrace("Error occured");
            LogTrace(ex.Message);
            LogTrace(ex.InnerException.Message);
        }

        /// <summary>
        /// This will appear on the Design Automation output
        /// </summary>
        private static void LogTrace(string format, params object[] args) { System.Console.WriteLine(format, args); }
    }
}
