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

        public virtual string Name
        {
            get
            {
                return StringResource.ResourceManager.GetString(this.NameResKey);
            }
        }
        public virtual string Description
        {
            get
            {
                return StringResource.ResourceManager.GetString(this.DescripResKey);
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

        public virtual string SortKey
        {
            get
            {
                return this.Name;
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

        protected virtual void Execute()
        {
            GrapeCity.Documents.Spread.Workbook workbook = new Spread.Workbook();
            this.Execute(workbook);
        }

        public virtual void Execute(GrapeCity.Documents.Spread.Workbook workbook)
        {

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
