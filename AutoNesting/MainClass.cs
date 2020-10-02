using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

using Autodesk.AutoCAD.Runtime;

namespace AutoNesting
{
    public class MainClass : IExtensionApplication
    {
        public void Initialize()
        {
            var cmds = helper.GetDllCmds();
            helper.AddCmdtoMenuBar(cmds, "排料工具");
        }

        public void Terminate()
        {
            //helper.DeleteCmdMenu( "排料工具");
        }
    }
}
