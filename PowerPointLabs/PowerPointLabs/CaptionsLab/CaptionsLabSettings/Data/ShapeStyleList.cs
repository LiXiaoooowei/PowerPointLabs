using System;
using System.Collections.ObjectModel;
using System.IO;
using System.Linq;
using System.Text;

using PowerPointLabs.ActionFramework.Common.Log;

namespace PowerPointLabs.CaptionsLab.CaptionsLabSettings.Data
{
    public class ShapeStyleList: ObservableCollection<ShapeStyle>
    {
        private DirectoryInfo _directory;

        public ShapeStyleList()
        {
        }

        public ShapeStyleList(string path) : this(new DirectoryInfo(path))
        {
        }

        public ShapeStyleList(DirectoryInfo directory)
        {
            _directory = directory;
            Update();
        }

        public string Path
        {
            set
            {
                _directory = new DirectoryInfo(value);
                Update();
            }
            get { return _directory.FullName; }
        }

        public DirectoryInfo DirectoryInfo
        {
            set
            {
                _directory = value;
                Update();
            }
            get { return _directory; }
        }

        private void Update()
        {
            if (Directory.Exists(Path))
            {
                foreach (var f in _directory.GetFiles("*.png"))
                {
                    Add(new ShapeStyle(f.FullName));
                }
            }
        }
    }
}
