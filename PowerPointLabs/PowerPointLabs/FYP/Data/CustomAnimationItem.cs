using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using Shape = Microsoft.Office.Interop.PowerPoint.Shape;

namespace PowerPointLabs.FYP.Data
{
    public class CustomAnimationItem: AnimationItem
    {
        public string ShapeName
        {
            get
            {
                return shape.Name;
            }
        }
        public string EffectName
        {
            get
            {
                return type.ToString();
            }
        }

        private Shape shape;
        private MsoAnimEffect type;
        private MsoAnimateByLevel level;
        private MsoTriState exit;

        public CustomAnimationItem(Shape shape, MsoAnimEffect type, MsoAnimateByLevel level, MsoTriState exit):base()
        {
            this.shape = shape;
            this.type = type;
            this.level = level;
            this.exit = exit;
        }

        public Shape GetShape()
        {
            return shape;
        }
        public MsoAnimEffect GetEffectType()
        {
            return type;
        }

        public MsoAnimateByLevel GetEffectLevel()
        {
            return level;
        }

        public MsoTriState GetExit()
        {
            return exit;
        }
    }
}
