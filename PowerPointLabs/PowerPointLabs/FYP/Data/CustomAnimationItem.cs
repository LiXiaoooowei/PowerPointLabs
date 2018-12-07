using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using Microsoft.Office.Interop.PowerPoint;

namespace PowerPointLabs.FYP.Data
{
    public class CustomAnimationItem: AnimationItem
    {
        public string ShapeName
        {
            get
            {
                return animatedEffect.Shape.Name;
            }
        }
        public string EffectName
        {
            get
            {
                return animatedEffect.EffectType.ToString();
            }
        }

        private Effect animatedEffect;

        public CustomAnimationItem(Effect effect)
        {
            animatedEffect = effect;
        }
    }
}
