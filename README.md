# Visio Wix Setup Project #

[Open project on Visual Studio Gallery](visualstudiogallery.msdn.microsoft.com/68d12e2d-eb42-4847-808a-7d80863bb90d)

The project template to create WiX-based installer for installing Microsoft Visio content (stencils and templates).

Installing (publishing) Microsoft Visio templates and stencils, in addition to copying files, requires also some additional registration to better degree of integration with Visio. This project adresses these issues.

Note that you need to have WiX Toolset installed in order to be able to use this template. 
You can get it from the official WiX toolset website: http://wixtoolset.org

The project creates ready-to-build project file which includes sample template and stencil (and publishing code for these). You may just add your own stencils/templates following this example.

The built installer should work for Visio 2007, 2010, 2013.

Both x32 and x64 versions are supported (two different build configurations)

For more information, visit [Unmanaged Visio](http://unmanagedvisio.com)


# Details #

The project includes a custom WiX Extension implementation, WiX Template, and VSIX project.

WiX Extension includes custom action, wix library, and the WiX compiler extension.