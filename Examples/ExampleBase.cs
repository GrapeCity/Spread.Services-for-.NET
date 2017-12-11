using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using GrapeCity.Documents.Spread;
using System.Reflection;
using System.Text.RegularExpressions;
using System.IO;

namespace GrapeCity.Documents.Spread.Examples
{
    public abstract class ExampleBase
    {
        public ExampleBase()
        {
        }

        public virtual string ID
        {
            get
            {
                return this.GetType().FullName;
            }
        }

        public string Code
        {
            get
            {
                return this.GetExampleCode();
            }
        }

        public virtual bool CanDownload
        {
            get
            {
                return true;
            }
        }

        public virtual bool ShowViewer
        {
            get
            {
                return true;
            }
        }


        public virtual bool ShowScreenshot
        {
            get
            {
                return false;
            }
        }


        public virtual bool ShowCode
        {
            get
            {
                return true;
            }
        }

        public virtual bool HasTemplate
        {
            get
            {
                return false;
            }
        }

        internal string UserAgent
        {
            get; set;
        }
        public virtual Stream GetTemplateStream(string templateName)
        {
            if (string.IsNullOrEmpty(templateName))
            {
                return null;
            }
            string resource = "GrapeCity.Documents.Spread.Examples.Resource.xlsx." + templateName;
            var assembly = this.GetType().GetTypeInfo().Assembly;
            return assembly.GetManifestResourceStream(resource);
        }

        public virtual string TemplateName
        {
            get
            {
                return null;
            }
        }

        public virtual bool IsViewReadOnly
        {
            get
            {
                return true;
            }
        }

        public virtual bool IsUpdate
        {
            get
            {
                return false;
            }
        }

        public virtual bool IsNew
        {
            get
            {
                return false;
            }
        }

        protected virtual string NameResKey
        {
            get
            {
                return this.GetType().Name + ".Name";
            }
        }

        protected virtual string DescripResKey
        {
            get
            {
                return this.GetType().Name + ".Descrip";
            }
        }

        protected string CurrentDirectory
        {
            get
            {
                return System.IO.Directory.GetCurrentDirectory();
            }
        }

        public void ExecuteExample(GrapeCity.Documents.Spread.Workbook workbook, string[] userAgents)
        {
            this.BeforeExecute(workbook, userAgents);
            this.Execute(workbook);
            this.AfterExecute(workbook, userAgents);
        }

        protected virtual void BeforeExecute(GrapeCity.Documents.Spread.Workbook workbook, string[] userAgents)
        {

        }

        public virtual void Execute(GrapeCity.Documents.Spread.Workbook workbook)
        {

        }

        protected virtual void AfterExecute(GrapeCity.Documents.Spread.Workbook workbook, string[] userAgents)
        {
            if (AgentIsMac(userAgents))
            {
                workbook.Calculate(); // ensure that all cached values can be saved in excel file, so number can display the file correctly even if the formulas are not supported in number.
            }
        }

        public virtual bool IsContainedInTree
        {
            get
            {
                return true;
            }
        }

        private string GetExampleCode()
        {
            string code = CodeResource.ResourceManager.GetString(this.GetType().FullName);
            if (!string.IsNullOrWhiteSpace(code))
            {
                code = Regex.Replace(code, "[\r\n][^\r\n]\\s{8}", "\n");
            }
            return code;
        }

        public string GetShortID()
        {
            return this.ID.Substring(this.ID.LastIndexOf(".") + 1);
        }

        public string ScreenshotBase64
        {
            get
            {
                if (ShowScreenshot)
                {
                    var id = GetType().FullName;
                    var assembly = GetType().GetTypeInfo().Assembly;
                    Stream stream = assembly.GetManifestResourceStream("GrapeCity.Documents.Spread.Examples.Resource.Screenshots." + id + ".png");
                    return ReadStreamToBase64(stream);
                }
                return null;
            }
        }

        public virtual string GetNameByCulture(string culture)
        {
            return StringResource.ResourceManager.GetString(this.NameResKey, new System.Globalization.CultureInfo(culture));
        }

        public virtual string GetDescriptionByCulture(string culture)
        {
            return StringResource.ResourceManager.GetString(this.DescripResKey, new System.Globalization.CultureInfo(culture));
        }

        protected bool AgentIsMac(string[] userAgents)
        {
            if (userAgents.Length > 0 && userAgents[0].ToLower().Contains("macintosh"))
            {
                return true;
            }
            return false;
        }

        private string ReadStreamToBase64(Stream input)
        {
            using (MemoryStream ms = new MemoryStream())
            {
                input.CopyTo(ms);
                return "data:image/png;base64," + Convert.ToBase64String(ms.ToArray());
            }
        }
    }


    public class FolderExample : ExampleBase
    {
        private List<ExampleBase> _children = null;
        private string _namespace;

        public FolderExample(string ns)
        {
            _namespace = ns;
        }

        public override string ID
        {
            get
            {
                return this._namespace;
            }
        }

        protected override string NameResKey
        {
            get
            {
                string shortName = _namespace.Substring(_namespace.LastIndexOf(".") + 1);
                return shortName + ".Name";
            }
        }

        protected override string DescripResKey
        {
            get
            {
                string shortName = _namespace.Substring(_namespace.LastIndexOf(".") + 1);
                return shortName + ".Descrip";
            }
        }

        public ExampleBase[] Children
        {
            get
            {
                if (_children == null)
                {
                    _children = this.GetChildren();
                }

                return _children.ToArray();
            }
        }

        private List<ExampleBase> GetChildren()
        {
            List<ExampleBase> children = new List<ExampleBase>();
            Type[] types = AssemblyUtility.GetTypesRecursively(_namespace);
            HashSet<string> subNS = new HashSet<string>();
            foreach (var type in types)
            {
                if (type.Namespace == _namespace)
                {
                    ExampleBase child = Activator.CreateInstance(type) as ExampleBase;
                    if (child.IsContainedInTree)
                    {
                        children.Add(child);
                    }
                }
                else if (!subNS.Contains(type.Namespace))
                {
                    string ends = type.Namespace.Substring(this._namespace.Length + 1);
                    if (!string.IsNullOrEmpty(ends))
                    {
                        var nsItems = ends.Split('.');
                        var currentNS = _namespace + "." + nsItems[0];
                        if (!subNS.Contains(currentNS))
                        {
                            children.Add(new FolderExample(currentNS));
                            subNS.Add(currentNS);
                        }
                        subNS.Add(type.Namespace);
                    }
                }
            }

            children.Sort(new ExampleComparer());

            return children;
        }

        public ExampleBase FindExample(string id)
        {
            return this.FindExample(this, id);
        }

        private ExampleBase FindExample(ExampleBase example, string id)
        {
            if (example.ID == id)
            {
                return example;
            }

            FolderExample folderExample = example as FolderExample;
            if (folderExample != null)
            {
                foreach (var child in folderExample.Children)
                {
                    ExampleBase result = this.FindExample(child, id);
                    if (result != null)
                    {
                        return result;
                    }
                }
            }

            return null;
        }

        public override bool IsNew
        {
            get
            {
                return false;
            }
        }

        public override bool IsUpdate
        {
            get
            {
                return IsUpdateRecursive(this);
            }
        }

        private bool IsUpdateRecursive(ExampleBase example)
        {
            if (example is FolderExample)
            {
                FolderExample childFolderExample = example as FolderExample;
                foreach (var item in childFolderExample.Children)
                {
                    if (item.IsUpdate || item.IsNew)
                    {
                        return true;
                    }

                    if (IsUpdateRecursive(item))
                    {
                        return true;
                    }
                }
            }
            else if (example.IsUpdate || example.IsNew)
            {
                return true;
            }

            return false;
        }
    }

    public static class AssemblyUtility
    {
        private static Assembly _assembly = null;
        private static List<Type> _types = null;
        private static Type _exampleBaseType = typeof(ExampleBase);
        private static Type _folderExampleType = typeof(FolderExample);

        static AssemblyUtility()
        {
            _assembly = typeof(Examples).GetTypeInfo().Assembly;
            _types = new List<Type>(_assembly.GetTypes());
            _types.Remove(_folderExampleType);
        }

        public static Type[] GetTypesRecursively(string ns)
        {
            return _types.FindAll(type => type.Namespace.StartsWith(ns) && type.GetTypeInfo().BaseType == _exampleBaseType).ToArray();
        }
    }
}
