using System;
using System.Runtime.InteropServices;
using System.ComponentModel.Design;
using System.Threading;
using System.Threading.Tasks;
using EnvDTE;
using Microsoft.VisualStudio;
using Microsoft.VisualStudio.Shell.Interop;

using Microsoft.VisualStudio.Shell;
using Task = System.Threading.Tasks.Task;

namespace GrzegorzKozub.VisualStudioExtensions.TotalCommanderLauncher
{
    [Guid(Guids.Package)]
    [InstalledProductRegistration("#1", "#2", "1.6.0.0", IconResourceID = 3)]
    [PackageRegistration(UseManagedResourcesOnly = true, AllowsBackgroundLoading = true)]
    [ProvideMenuResource("Menus.ctmenu", 1)]
    [ProvideOptionPage(typeof(Options), "Total Commander Launcher", "General", 0, 0, false)]
    [ProvideAutoLoad(VSConstants.UICONTEXT.SolutionExistsAndFullyLoaded_string, PackageAutoLoadFlags.BackgroundLoad)]
    public sealed class TotalCommanderLauncherPackage : AsyncPackage
    {
        private IMenuCommandService _menuCommandService;
        private IVsUIShell _uiShell;

        #region Package Members

        protected override async Task InitializeAsync(CancellationToken cancellationToken, IProgress<ServiceProgressData> progress)
        {
            await base.InitializeAsync(cancellationToken, progress);

            await JoinableTaskFactory.SwitchToMainThreadAsync(cancellationToken);

            _menuCommandService = await GetServiceAsync(typeof(IMenuCommandService)) as OleMenuCommandService;
            _uiShell            = await GetServiceAsync(typeof(SVsUIShell)) as IVsUIShell;

            AddMenuCommand(CommandIds.TotalCommander, HandleTotalCommanderMenuCommand, options => true);
        }

        #endregion

        private void AddMenuCommand(uint commandId, EventHandler invokeHandler, Func<Options, bool> visible)
        {
            var commandID = new CommandID(Guids.MenuGroup, (int)commandId);
            var menuCommand = new OleMenuCommand(invokeHandler, commandID);

            menuCommand.BeforeQueryStatus += (s, e) =>
            {
                menuCommand.Visible = visible(GetDialogPage(typeof(Options)) as Options);
            };

            _menuCommandService.AddCommand(menuCommand);
        }

        private void HandleTotalCommanderMenuCommand(object sender, EventArgs eventArgs)
        {
            var dte = GetGlobalService(typeof(DTE)) as DTE;
            var options = GetDialogPage(typeof(Options)) as Options;

            var validationErrors = options.GetValidationErrors();

            if (validationErrors != null)
            {
                DisplayErrorAndSuggestOptions(validationErrors);
                return;
            }

            var launcher = new Launcher(dte, options);

            try
            {
                launcher.Launch();
            }
            catch (Exception e)
            {
                DisplayErrorAndSuggestOptions(e.Message);
            }
        }

        private void DisplayErrorAndSuggestOptions(string errorMessage)
        {
            var comp = Guid.Empty;
            int result;

            _uiShell.ShowMessageBox(
                0,
                ref comp,
                errorMessage,
                "Do you want to visit the Options page for Total Commander Launcher now?",
                string.Empty,
                0,
                OLEMSGBUTTON.OLEMSGBUTTON_YESNO,
                OLEMSGDEFBUTTON.OLEMSGDEFBUTTON_FIRST,
                OLEMSGICON.OLEMSGICON_CRITICAL,
                0,
                out result);

            if (result == 6)
            {
                var optionsCommandId = new CommandID(VSConstants.GUID_VSStandardCommandSet97, VSConstants.cmdidToolsOptions);
                ((MenuCommandService)_menuCommandService).GlobalInvoke(optionsCommandId, Guids.Options);
            }
        }
    }
}
