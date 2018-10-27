using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using Microsoft.Office.Interop.PowerPoint;

using PowerPointLabs.ActionFramework.Common.Log;
using PowerPointLabs.Models;
using PowerPointLabs.Tags;

using Shape = Microsoft.Office.Interop.PowerPoint.Shape;

namespace PowerPointLabs.CaptionsLab
{
    public static class AnimationUtil
    {
        public static void UpdateAnimationsForCalloutsOnSlide(IntermediateResultTable intermediateResult, PowerPointSlide slide)
        {
            List<Tuple<NameTag, string>> notesInserted = intermediateResult.GetInsertedNotes();
            List<Tuple<NameTag, string>> notesDeleted = intermediateResult.GetDeletedNotes();
            Sequence sequence = slide.TimeLine.MainSequence;
            IEnumerable<Effect> mainEffects = sequence.Cast<Effect>();
            // handle deleted notes
            for (int i = 0; i < notesDeleted.Count; i++)
            {
                List<Shape> shapes = slide.GetShapeWithName("PPTLabs Callout " + notesDeleted[i].Item1.Contents);
                slide.RemoveAnimationsForShapes(shapes);
            }
            // handle reordered notes
            int prevClick = -1;
            int nextClick = 0;
            HashSet<Shape> shapeSet = new HashSet<Shape>();
            List<Tuple<Shape, MsoAnimEffect, int>> shapesOrder = new List<Tuple<Shape, MsoAnimEffect, int>>();

            foreach (Effect animeEffect in mainEffects)
            {
                Effect effectForClick = sequence.FindFirstAnimationForClick(nextClick);
                if (effectForClick == null)
                {
                    Logger.Log("effect is null");
                }
                else
                {
                    Logger.Log("click is " + nextClick + effectForClick.Shape.Name);
                    Logger.Log("click is " + nextClick + animeEffect.Shape.Name);
                }
                if (nextClick == 0 && effectForClick == null)
                {
                    effectForClick = sequence.FindFirstAnimationForClick(1);
                    nextClick++;
                    prevClick++;
                }
                if (effectForClick != null && animeEffect.Shape.Name == effectForClick.Shape.Name)
                {
                    Logger.Log("we reached this place");
                    nextClick++;
                    prevClick++;
                }
                Shape shape = animeEffect.Shape;
                if (!shapeSet.Contains(shape))
                {
                    if (!shape.Name.Contains("PPTLabs Callout") || animeEffect.Exit != Microsoft.Office.Core.MsoTriState.msoTrue)
                    {
                        Logger.Log("adding " + shape.Name + " with click " + prevClick);
                        shapeSet.Add(shape);
                        shapesOrder.Add(new Tuple<Shape, MsoAnimEffect, int>(animeEffect.Shape, animeEffect.EffectType, prevClick));
                    }
                }
            }
            slide.RemoveAnimationsForShapes(shapeSet.ToList());

            Shape prevShape = null;
            
            foreach (Tuple<Shape, MsoAnimEffect, int> shapeEffect in shapesOrder)
            {
                Shape shape = shapeEffect.Item1;
                MsoAnimEffect originalEffect = shapeEffect.Item2;
                int animeClick = shapeEffect.Item3;
                Logger.Log("shape " + shape.Name + " click " + animeClick);
                Effect newEffect;
                if (animeClick == 0)
                {
                    newEffect = slide.SetShapeAsAutoplay(shape);
                }
                else
                {
                    newEffect = slide.ShowShapeAfterClick(shape, animeClick);
                }
                newEffect.EffectType = originalEffect;
                if (shape.Name.Contains("PPTLabs Callout "))
                {
                    if (prevShape != null)
                    {
                        slide.HideShapeAfterClick(prevShape, animeClick);
                    }
                    prevShape = shape;
                }
            }
            
            // handle newly inserted notes
            int clickNo = 0;
            Effect effect = sequence.FindFirstAnimationForClick(clickNo++);
            if (effect == null)
            {
                effect = sequence.FindFirstAnimationForClick(clickNo);
            }
            while (effect != null)
            {
                effect = sequence.FindFirstAnimationForClick(++clickNo);
            }           
            for (int i = 0; i < notesInserted.Count; i++)
            {
                List<Shape> shapes = slide.GetShapeWithName("PPTLabs Callout " + notesInserted[i].Item1.Contents);
                if (shapes.Count == 0)
                {
                    continue;
                }
                Shape currShape = shapes[0];
                Effect showEffect = slide.ShowShapeAfterClick(currShape, clickNo);
                if (prevShape != null)
                {
                    slide.HideShapeAfterClick(prevShape, clickNo);
                }
                clickNo++;
                prevShape = currShape;
            }
        }
    }
}
