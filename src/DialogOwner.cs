#region License statement
/* SnakeTail is free software: you can redistribute it and/or modify
 * it under the terms of the GNU General Public License as published by
 * the Free Software Foundation, version 3 of the License.
 *
 * This program is distributed in the hope that it will be useful,
 * but WITHOUT ANY WARRANTY; without even the implied warranty of
 * MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
 * GNU General Public License for more details.
 *
 * You should have received a copy of the GNU General Public License
 * along with this program.  If not, see <http://www.gnu.org/licenses/>.
 */
#endregion

using System;
using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace SnakeTail
{
    /// <summary>统一解析对话框 owner，避免弹窗失去父窗口。</summary>
    static class DialogOwner
    {
        [DllImport("user32.dll")]
        private static extern IntPtr GetForegroundWindow();

        private sealed class WindowHandleOwner : IWin32Window
        {
            public IntPtr Handle { get; private set; }

            public WindowHandleOwner(IntPtr handle)
            {
                Handle = handle;
            }
        }

        public static IWin32Window Resolve()
        {
            // 优先主窗口，确保弹窗跟随应用主界面
            if (MainForm.Instance != null && !MainForm.Instance.IsDisposed && MainForm.Instance.IsHandleCreated)
                return MainForm.Instance;

            // 次选当前活动窗体，兼容子窗口场景
            Form activeForm = Form.ActiveForm;
            if (activeForm != null && !activeForm.IsDisposed && activeForm.IsHandleCreated)
                return activeForm;

            // 最后回退前台窗口句柄，避免无 owner 的顶层弹窗
            IntPtr foreground = GetForegroundWindow();
            if (foreground != IntPtr.Zero)
                return new WindowHandleOwner(foreground);

            return null;
        }
    }
}
