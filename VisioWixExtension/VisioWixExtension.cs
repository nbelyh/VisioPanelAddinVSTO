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
    }
}
