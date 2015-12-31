using System.Reflection;
using Microsoft.Tools.WindowsInstallerXml;

namespace VisioWixExtension
{
    public class VisioWixExtension : WixExtension 
    {
        private VisioCompilerExtension _compilerExtensionExtension;

        /// <summary>
        /// Gets the optional compiler extension.
        /// </summary>
        /// <value>The optional compiler extension.</value>
        public override CompilerExtension CompilerExtension
        {
            get
            {
                if (null == _compilerExtensionExtension)
                    _compilerExtensionExtension = new VisioCompilerExtension();

                return _compilerExtensionExtension;
            }
        }

        private VisioDecompilerExtension _decompilerExtension;

        /// <summary>
        /// Gets the optional decompiler extension.
        /// </summary>
        /// <value>The optional decompiler extension.</value>
        public override DecompilerExtension DecompilerExtension
        {
            get
            {
                if (null == _decompilerExtension)
                {
                    _decompilerExtension = new VisioDecompilerExtension();
                }

                return _decompilerExtension;
            }
        }

        private VisioBinderExtension _binderExtension;

        /// <summary>
        /// Gets the optional binder extension.
        /// </summary>
        /// <value>The optional binder extension.</value>
        public override BinderExtension BinderExtension
        {
            get
            {
                if (null == _binderExtension)
                {
                    _binderExtension = new VisioBinderExtension();
                }

                return _binderExtension;
            }
        }

        private TableDefinitionCollection _tableDefinitions;

        public override TableDefinitionCollection TableDefinitions
        {
            get {
                if (null == _tableDefinitions)
                    _tableDefinitions = LoadTableDefinitionHelper(Assembly.GetExecutingAssembly(), "VisioWixExtension.Data.Tables.xml");

                return _tableDefinitions;
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
