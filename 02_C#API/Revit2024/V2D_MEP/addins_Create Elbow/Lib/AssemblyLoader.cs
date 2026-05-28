using System;
using System.Collections.Generic;
using System.IO;
using System.Reflection;

namespace AlphaBIM
{
    /// <summary>
    /// Tự động load tất cả DLL phụ thuộc trong cùng thư mục với addin.
    /// Giúp Revit tìm thấy các assembly được tham chiếu khi chạy lệnh.
    /// </summary>
    internal class AssemblyLoader : IDisposable
    {
        private static string ExecutingPath => Assembly.GetExecutingAssembly().Location;

        internal AssemblyLoader()
        {
            AppDomain.CurrentDomain.AssemblyResolve += LoadFromSameFolder;
        }

        private static Assembly LoadFromSameFolder(object sender, ResolveEventArgs args)
        {
            if (string.IsNullOrEmpty(ExecutingPath)) return null;

            string GetAssemblyName(string fullName) =>
                fullName.Substring(0, fullName.IndexOf(','));

            var dir = new FileInfo(ExecutingPath).Directory;
            if (dir == null) return null;

            foreach (var file in dir.EnumerateFiles())
            {
                if (!file.Name.EndsWith(".dll") && !file.Name.EndsWith(".exe"))
                    continue;

                try
                {
                    var assembly = Assembly.LoadFrom(file.FullName);
                    if (GetAssemblyName(assembly.FullName) == GetAssemblyName(args.Name))
                        return assembly;
                }
                catch
                {
                    // Bỏ qua lỗi load từng file đơn lẻ
                }
            }

            return null;
        }

        void IDisposable.Dispose()
        {
            AppDomain.CurrentDomain.AssemblyResolve -= LoadFromSameFolder;
        }

        /// <summary>
        /// Load tất cả DLL trong thư mục addin vào AppDomain hiện tại.
        /// Gọi một lần khi Execute() bắt đầu.
        /// </summary>
        internal static List<Assembly> LoadAllRibbonAssemblies(string commandFolderPath)
        {
            var list = new List<Assembly>();

            if (string.IsNullOrEmpty(commandFolderPath) ||
                !Directory.Exists(commandFolderPath))
                return list;

            foreach (string assemblyFile in Directory.GetFiles(commandFolderPath, "*.dll"))
            {
                try
                {
                    list.Add(Assembly.LoadFrom(assemblyFile));
                }
                catch
                {
                    // Bỏ qua DLL không load được (VD: native DLL)
                }
            }

            return list;
        }
    }
}
