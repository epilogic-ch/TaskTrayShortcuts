using System;
using System.Diagnostics;
using System.IO;
using System.Collections.Generic;
using System.Windows.Forms;
using System.ComponentModel;
using Shell32;
using System.Text;
using System.Runtime.InteropServices;
using System.Drawing;

namespace TaskTrayShortcuts
{
    public class TaskTrayShortcutsContext : ApplicationContext
    {
        [DllImport("Shell32.dll", EntryPoint = "ExtractIconExW", CharSet = CharSet.Unicode, ExactSpelling = true, CallingConvention = CallingConvention.StdCall)]
        public static extern int ExtractIconEx(string sFile, int iIndex, out IntPtr piLargeVersion, out IntPtr piSmallVersion, int amountIcons);

        string folderPath = null;
        NotifyIcon notifyIcon = new NotifyIcon();
        Configuration configWindow = new Configuration();

        public TaskTrayShortcutsContext(string[] args)
        {
            #if DEBUG
                this.folderPath = @"D:\_";
            #else
                // We expect a folder path as first argument
                if (args.Length == 0)
                {                    
                    MessageBox.Show("Missing parameter. Please add a path as parameter.\n\n"+
                        "Syntax: TaskTrayShortcuts.exe <path>\n"+
                        "where <path> is the path containing shortcuts");
                    System.Environment.Exit(1);
                }

                this.folderPath = args[0];
                if (!Directory.Exists(this.folderPath))
                {
                    MessageBox.Show("Path does not exists.");
                    System.Environment.Exit(1);
                }
            #endif


            List<ToolStripMenuItem> items = this.ProcessDirectory(this.folderPath, true);
            System.Drawing.Icon test = IconExtractor.Extract("shell32.dll", 131, true);
            items.Add(new ToolStripMenuItem("Exit", test.ToBitmap(), new EventHandler(Exit)));

            ContextMenuStrip menu = new ContextMenuStrip();
            menu.Items.AddRange(items.ToArray());

            notifyIcon.Icon = TaskTrayShortcuts.Properties.Resources.AppIcon;
            notifyIcon.DoubleClick += new EventHandler(Open);
            notifyIcon.Click += new EventHandler(Open);
            notifyIcon.ContextMenuStrip = menu;
            notifyIcon.Visible = true;
        }

        // Process all files in the directory passed in, recurse on any directories
        // that are found, and process the files they contain.
        public List<ToolStripMenuItem> ProcessDirectory(string targetDirectory, bool addOpenFolderLink)
        {
            List<ToolStripMenuItem> menu = new List<ToolStripMenuItem>();
            ToolStripMenuItem item;
            FileInfo fInfo;

            System.Drawing.Bitmap folderImage = IconExtractor.Extract("shell32.dll", 3, true).ToBitmap();

            try
            {
                if (addOpenFolderLink)
                {
                    fInfo = new FileInfo(targetDirectory);
                    item = new ToolStripMenuItem("Open folder \"" + fInfo.Name + "\"", null, delegate (object sender, EventArgs e)
                    {
                        this.ExecuteCommand(sender, e, targetDirectory);
                    });
                    item.Image = IconExtractor.Extract("shell32.dll", 3, true).ToBitmap();
                    menu.Add(item);
                }

                // Recurse into subdirectories of this directory.
                string[] subdirectoryEntries = Directory.GetDirectories(targetDirectory);
                foreach (string subdirectory in subdirectoryEntries)
                {
                    fInfo = new FileInfo(subdirectory);
                    if (!fInfo.Attributes.HasFlag(FileAttributes.Hidden))
                    {
                        item = new ToolStripMenuItem(fInfo.Name, null, ProcessDirectory(subdirectory, true).ToArray());
                        item.Image = folderImage;
                        menu.Add(item);
                    }
                }

                // Process the list of files found in the directory.
                string[] fileEntries = Directory.GetFiles(targetDirectory);
                foreach (string fileName in fileEntries)
                {
                    if (!fileName.ToLower().StartsWith("$"))
                    {
                        fInfo = new FileInfo(fileName);
                        if (!fInfo.Attributes.HasFlag(FileAttributes.Hidden))
                        {
                            // Extract shortcut name and icon
                            string name = fInfo.Name.Substring(0, fInfo.Name.LastIndexOf(fInfo.Extension));

                            item = new ToolStripMenuItem(name, null, delegate (object sender, EventArgs e) {
                                this.ExecuteCommand(sender, e, fileName);
                            });

                            try
                            {
                                item.Image = GetShortcutTargetIcon(fileName);
                            }
                            catch (System.IO.FileNotFoundException) { }

                            menu.Add(item);
                        }
                    }
                }
            } catch (DirectoryNotFoundException) { }

            return menu;
        }

        public static System.Drawing.Bitmap GetShortcutTargetIcon(string shortcutFilename)
        {
            System.Drawing.Icon icon;

            string pathOnly = Path.GetDirectoryName(shortcutFilename);
            string filenameOnly = Path.GetFileName(shortcutFilename);
            string target = shortcutFilename;

            try
            {
                Shell shell = new Shell();
                Folder folder = shell.NameSpace(pathOnly);
                FolderItem folderItem = folder.ParseName(filenameOnly);
                if (folderItem != null)
                {
                    string tmp;
                    ShellLinkObject link = (ShellLinkObject) folderItem.GetLink;

                    // Get target
                    if (link != null && link.Path != string.Empty)
                    {
                        target = link.Path;
                    }

                    // Get icon
                    int val = link.GetIconLocation(out tmp);

                    IntPtr largeIconPtr = IntPtr.Zero;
                    IntPtr smallIconPtr = IntPtr.Zero;
                    ExtractIconEx(tmp, val, out largeIconPtr, out smallIconPtr, 1);
                    if (smallIconPtr != IntPtr.Zero)
                    {
                        icon = Icon.FromHandle(smallIconPtr);
                        if (icon != null && icon.Width == 16 && icon.Height == 16)
                        {
                            return icon.ToBitmap();
                        }
                    }
                }
            }
            catch (System.NotImplementedException) { }
            catch (System.ArgumentException) { }

            // Alternate method for getting icon
            icon = IconReader.GetFileIcon(target, IconReader.IconSize.Small, false);
            Console.WriteLine(target + "\t" + icon.ToString());
            if (icon != null && icon.Width == 16 && icon.Height == 16) return icon.ToBitmap();

            icon = IconReader.GetFileIcon(shortcutFilename, IconReader.IconSize.Small, false);
            Console.WriteLine(shortcutFilename + "\t" + icon.ToString());
            if (icon != null && icon.Width == 16 && icon.Height == 16) return icon.ToBitmap();

            icon = System.Drawing.Icon.ExtractAssociatedIcon(target);
            Console.WriteLine(target + "\t" + icon.ToString());
            if (icon != null && icon.Width == 16 && icon.Height == 16) return resizeIcon(icon);

            return resizeIcon(System.Drawing.Icon.ExtractAssociatedIcon(shortcutFilename));
        }

        public static System.Drawing.Bitmap resizeIcon(System.Drawing.Icon icon)
        {
            if (icon != null)
            {
                System.Drawing.Bitmap destinationImage = new System.Drawing.Bitmap(16, 16);
                System.Drawing.Bitmap bitmap = new System.Drawing.Bitmap(icon.ToBitmap());
                System.Drawing.Rectangle destinationRect = new System.Drawing.Rectangle(0, 0, 16, 16);
                using (System.Drawing.Graphics graphics = System.Drawing.Graphics.FromImage(destinationImage))
                {
                    //graphics.CompositingMode = System.Drawing.Drawing2D.CompositingMode.SourceCopy;
                    //graphics.CompositingQuality = System.Drawing.Drawing2D.CompositingQuality.HighQuality;
                    graphics.InterpolationMode = System.Drawing.Drawing2D.InterpolationMode.HighQualityBicubic;
                    graphics.DrawImage(bitmap, destinationRect, 0, 0, bitmap.Width, bitmap.Height, System.Drawing.GraphicsUnit.Pixel);
                }

                return destinationImage;
            }
            return null;
        }

        void ExecuteCommand(object sender, EventArgs e, string fileName)
        {
            try
            {
                Process proc = new Process();
                proc.StartInfo.FileName = fileName;
                proc.Start();
            }
            catch (Win32Exception)
            {
                MessageBox.Show("Unable to start \"" + fileName + "\" process");
            }
        }

        void Open(object sender, EventArgs e)
        {
            MouseEventArgs me = (MouseEventArgs)e;
            if (me.Button is MouseButtons.Left)
            {
                if (notifyIcon.ContextMenuStrip.Visible)
                {
                    notifyIcon.ContextMenuStrip.Hide();
                }
                else
                {
                    System.Drawing.Point pos = Cursor.Position;
                    pos.X -= notifyIcon.ContextMenuStrip.Width;
                    pos.Y -= notifyIcon.ContextMenuStrip.Height;
                    if (pos.X < 0) pos.X = 0;
                    if (pos.Y < 0) pos.Y = 0;
                    notifyIcon.ContextMenuStrip.Show(pos);
                }
            }
        }

        void Exit(object sender, EventArgs e)
        {
            // We must manually tidy up and remove the icon before we exit.
            // Otherwise it will be left behind until the user mouses over.
            notifyIcon.Visible = false;
            Application.Exit();
        }
    }
}
