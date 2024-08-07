using Bentley.MstnPlatformNET;
using System.Windows.Forms;

namespace csAddins
{
    [AddIn(MdlTaskID = "csAddins")]
    //每个Addin项目必须且只包含一个从Bentley.MstnPlatformNET.AddIn派生下来的class，同时必须实现其纯虚成员函数Run
    internal sealed class MyAddin : AddIn
    {
        public static MyAddin Addin = null;

        //此构造函数是我们的程序集加载时第一个被执行的函数
        private MyAddin(System.IntPtr mdlDesc) : base(mdlDesc)
        {
            Addin = this;
        }

        //Run函数是构造函数执行完以后第一个被执行的函数，有点类似与NativeCode项目中的Main函数
        protected override int Run(string[] commandLine)
        {
            return 0;
        }
    }
}
