//------------------------------------------------------------------------------
// <copyright file="OpenBinFolderPackage.cs" company="Company">
//     Copyright (c) Company.  All rights reserved.
// </copyright>
//------------------------------------------------------------------------------

using System;
using System.ComponentModel.Design;
using System.Diagnostics;
using System.Diagnostics.CodeAnalysis;
using System.Globalization;
using System.Runtime.InteropServices;
using Microsoft.VisualStudio;
using Microsoft.VisualStudio.OLE.Interop;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using Microsoft.Win32;
using EnvDTE80;
using EnvDTE;
using System.IO;

namespace OpenBinFolder
{
   
    [PackageRegistration(UseManagedResourcesOnly = true)]
    [InstalledProductRegistration("#110", "#112", "1.0")] // Info on this package for Help/About
    [Guid(OpenBinFolderPackage.PackageGuidString)]
    [SuppressMessage("StyleCop.CSharp.DocumentationRules", "SA1650:ElementDocumentationMustBeSpelledCorrectly", Justification = "pkgdef, VS and vsixmanifest are valid VS terms")]
    [ProvideMenuResource("Menus.ctmenu", 1)]
    public sealed class OpenBinFolderPackage : Package
    {
        /// <summary>
        /// OpenBinFolderPackage GUID string.
        /// </summary>
        public const string PackageGuidString = "21572e2d-a591-4cd1-b073-c4ae5e3f6be6";

        /// <summary>
        /// Initializes a new instance of the <see cref="OpenBinFolderPackage"/> class.
        /// </summary>
        public OpenBinFolderPackage()
        {
            
        }

        #region Package Members

        /// <summary>
        /// Initialization of the package; this method is called right after the package is sited, so this is the place
        /// where you can put all the initialization code that rely on services provided by VisualStudio.
        /// </summary>
        protected override void Initialize()
        {
            base.Initialize();

            //get the menu service
            OleMenuCommandService _menuCommandService = GetService(typeof(IMenuCommandService)) as OleMenuCommandService;

            //create the click method for the menu command we are adding to the right click of the project menu
            CommandID _openBinFolderCommand = new CommandID(Guid.Parse("{02AB237F-F580-4278-A02B-8DA88483528E}"), int.Parse("3B9ACA01", System.Globalization.NumberStyles.HexNumber));
            MenuCommand _openBinFolderMenu = new MenuCommand(OpenBinFolderWithFileExplorer, _openBinFolderCommand);
            _menuCommandService.AddCommand(_openBinFolderMenu);
        }

        private void OpenBinFolderWithFileExplorer(object sender, EventArgs e)
        {
            //grab the DTE object
            var dte = (DTE2)this.GetService(typeof(DTE));
            //Get the active projects within the solution.
            Array _activeProjects = (Array)dte.ActiveSolutionProjects;

            //loop through each active project
            foreach (Project _activeProject in _activeProjects)
            {
                //get the directory path based on the project file.
                string _projectPath = Path.GetDirectoryName(_activeProject.FullName);
                //combine the bin directory to the project directory
                string _projectBinPath = Path.Combine(_projectPath, "bin");
                //get the active configuration name
                string _configName = _activeProject.ConfigurationManager.ActiveConfiguration.ConfigurationName;

                //combine the active config to the directory bin location
                string _projectBinBuildPath = Path.Combine(_projectBinPath, _configName);

                //if the directory exists (already built) then open that directory
                //in windows explorer using the diagnostics.process object
                if (Directory.Exists(_projectBinBuildPath))
                {
                    System.Diagnostics.Process.Start(_projectBinBuildPath);
                }
                else
                {
                    //if the directory doesnt exist, the open the bin directory (one level up).
                    System.Diagnostics.Process.Start(_projectBinPath);
                }
            }
        }
        #endregion
    }
}
