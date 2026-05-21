using System;
using System.Runtime.InteropServices;
using System.Threading;

class Click {
    [DllImport("user32.dll")] static extern bool SetCursorPos(int x, int y);
    [DllImport("user32.dll")] static extern void mouse_event(uint flags, uint x, uint y, uint data, IntPtr extra);
    const uint MOUSEEVENTF_LEFTDOWN = 0x0002, MOUSEEVENTF_LEFTUP = 0x0004;

    static int Main(string[] args) {
        if (args.Length < 2) { Console.WriteLine("usage: Click X Y"); return 2; }
        int x = int.Parse(args[0]);
        int y = int.Parse(args[1]);
        SetCursorPos(x, y);
        Thread.Sleep(150);
        mouse_event(MOUSEEVENTF_LEFTDOWN, 0, 0, 0, IntPtr.Zero);
        Thread.Sleep(60);
        mouse_event(MOUSEEVENTF_LEFTUP, 0, 0, 0, IntPtr.Zero);
        Console.WriteLine("click (" + x + "," + y + ") done");
        return 0;
    }
}
