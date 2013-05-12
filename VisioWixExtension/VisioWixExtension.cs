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
                if (null == this._compilerExtensionExtension)
                {
                    this._compilerExtensionExtension = new VisioCompilerExtension();
                }

                return this._compilerExtensionExtension;
            }
        }
    }
}
