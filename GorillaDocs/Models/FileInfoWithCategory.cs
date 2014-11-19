using System;
using System.IO;
using System.Runtime.Serialization;
using System.Windows.Media.Imaging;

namespace GorillaDocs.Models
{
    [DataContract]
    public class FileWithCategory
    {
        string _Category;
        FileInfo file;
        public FileWithCategory() { /* Parameterless constructor only used for Serialization*/ }
        public FileWithCategory(FileInfo file) { this.file = file; }
        public FileWithCategory(string file) { this.file = new FileInfo(file); }

        public string Category
        {
            get
            {
                if (string.IsNullOrEmpty(_Category))
                    _Category = file.Path().Substring(file.Path().LastIndexOf("\\") + 1);
                return _Category;
            }
            set { _Category = value; }
        }
        [DataMember]
        public string FullName { get { return file.FullName; } set { file = new FileInfo(value); } }
        public string Name { get { return file.Name; } }
        public string NameWithoutExtension { get { return file.NameWithoutExtension(); } }
        public BitmapFrame Image
        {
            get
            {
                if (file.IsWord())
                    return Properties.Resources.WordTemplate.AsBitmapFrame();
                else if (file.IsPowerPoint())
                    return Properties.Resources.PowerPointTemplate.AsBitmapFrame();
                else
                    throw new InvalidOperationException("Unknown file type.");
            }
        }
    }
}
