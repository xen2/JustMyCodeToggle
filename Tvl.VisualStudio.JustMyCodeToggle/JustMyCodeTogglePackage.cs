// Copyright (c) Tunnel Vision Laboratories, LLC. All Rights Reserved.
// Licensed under the MIT License. See LICENSE.txt in the project root for license information.

namespace Tvl.VisualStudio.JustMyCodeToggle
{
    using System;
    using System.ComponentModel.Design;
    using System.Runtime.InteropServices;
    using System.Threading;
    using Microsoft;
    using Microsoft.VisualStudio;
    using Microsoft.VisualStudio.Shell;
    using IMenuCommandService = System.ComponentModel.Design.IMenuCommandService;
    using Task = System.Threading.Tasks.Task;

    [Guid(JustMyCodeToggleConstants.GuidJustMyCodeTogglePackageString)]
    [PackageRegistration(UseManagedResourcesOnly = true, AllowsBackgroundLoading = true)]
    [ProvideMenuResource(1000, 1)]
    [ProvideAutoLoad(VSConstants.UICONTEXT.SolutionExistsAndFullyLoaded_string, PackageAutoLoadFlags.BackgroundLoad)]
    internal class JustMyCodeTogglePackage : AsyncPackage
    {
        private readonly OleMenuCommand _justMyCodeCommand;
        private readonly OleMenuCommand _diagnosticBuildLogCommand;

        public JustMyCodeTogglePackage()
        {
            var justMyCodeCommandId = new CommandID(JustMyCodeToggleConstants.GuidJustMyCodeToggleCommandSet, JustMyCodeToggleConstants.CmdidJustMyCodeToggle);
            _justMyCodeCommand = SetupCommand(justMyCodeCommandId, "Debugging", "General", "EnableJustMyCode",
                value => value is bool b && !b,
                value => value is bool b && b);

            var diagnosticBuildLogCommandId = new CommandID(JustMyCodeToggleConstants.GuidJustMyCodeToggleCommandSet, JustMyCodeToggleConstants.CmdidDiagnosticBuildLogToggle);
            _diagnosticBuildLogCommand = SetupCommand(diagnosticBuildLogCommandId, "Environment", "ProjectsAndSolution", "MSBuildOutputVerbosity",
                value => value is int i && i == 4 ? 1 : 4,
                value => value is int i && i == 4);
        }

        public EnvDTE.DTE ApplicationObject
        {
            get
            {
                return GetService(typeof(EnvDTE._DTE)) as EnvDTE.DTE;
            }
        }

        protected override async Task InitializeAsync(CancellationToken cancellationToken, IProgress<ServiceProgressData> progress)
        {
            await base.InitializeAsync(cancellationToken, progress);

            await JoinableTaskFactory.SwitchToMainThreadAsync(cancellationToken);

            var mcs = (IMenuCommandService)await GetServiceAsync(typeof(IMenuCommandService));
            Assumes.Present(mcs);
            mcs.AddCommand(_justMyCodeCommand);
            mcs.AddCommand(_diagnosticBuildLogCommand);
        }

        private OleMenuCommand SetupCommand(CommandID id, string category, string page, string item, Func<object, object> toggleValue, Func<object, bool> isChecked)
        {
            EventHandler invokeHandler = (sender, e) =>
            {
                try
                {
                    EnvDTE.Property property = ApplicationObject.get_Properties(category, page).Item(item);
                    property.Value = toggleValue(property.Value);
                }
                catch (Exception ex) when (!ErrorHandler.IsCriticalException(ex))
                {
                }
            };
            EventHandler changeHandler = (sender, e) => { };
            EventHandler beforeQueryStatus = (sender, e) =>
            {
                var command = sender as OleMenuCommand;
                try
                {
                    command.Supported = true;

                    EnvDTE.Property property = ApplicationObject.get_Properties(category, page).Item(item);
                    command.Checked = isChecked(property.Value);
                    command.Enabled = true;
                }
                catch (Exception ex) when (!ErrorHandler.IsCriticalException(ex))
                {
                    command.Supported = false;
                    command.Enabled = false;
                }
            };
            return new OleMenuCommand(invokeHandler, changeHandler, beforeQueryStatus, id);
        }
    }
}
