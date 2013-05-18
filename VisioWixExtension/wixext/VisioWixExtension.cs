using System.Reflection;
using Microsoft.Tools.WindowsInstallerXml;

namespace VisioWixExtension
{
    public class VisioWixExtension : WixExtension 
    {
        private RowGenerator _compilerExtensionExtension;

        /// <summary>
        /// Gets the optional compiler extension.
        /// </summary>
        /// <value>The optional compiler extension.</value>
        public override CompilerExtension CompilerExtension
        {
            get
            {
                if (null == _compilerExtensionExtension)
                    _compilerExtensionExtension = new RowGenerator();

                return _compilerExtensionExtension;
            }
        }

        private Library _library;
        /// <summary>
        /// Gets the library associated with this extension.
        /// </summary>
        /// <param name="tableDefinitions">The table definitions to use while loading the library.</param>
        /// <returns>The loaded library.</returns>
        public override Library GetLibrary(TableDefinitionCollection tableDefinitions)
        {
            if (null == _library)
                _library = LoadLibraryHelper(Assembly.GetExecutingAssembly(), "VisioWixExtension.Data.Visio.wixlib", tableDefinitions);

            return _library;
        }

    }
}
