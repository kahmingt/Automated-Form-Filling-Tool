using System.Runtime.CompilerServices;

namespace Utilities
{
    /// <summary>
    /// <para>
    /// Get the root folder of a C# project is to leverage [CallerFilePath] attribute to obtain the
    /// full path name of a source file, and then subtract the filename plus extension from it,
    /// leaving you with the path to the project.
    /// </para>
    /// https://stackoverflow.com/questions/816566/how-do-you-get-the-current-project-directory-from-c-sharp-code-when-creating-a-c#66285728
    /// </summary>
    internal static class ProjectSourcePath
    {
        private const string myRelativePath = nameof(ProjectSourcePath) + ".cs";
        private static string? lazyValue;
        public static string Value => lazyValue ??= calculatePath();

        private static string calculatePath()
        {
            string pathName = GetSourceFilePathName();
            return pathName.Substring(0, pathName.Length - myRelativePath.Length);
        }
    
        public static string GetSourceFilePathName([CallerFilePath] string? callerFilePath = null) => callerFilePath ?? "";
    }
}

