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
            IEnumerable<NameTag> notes = intermediateResult.GetNotes();

            Sequence sequence = slide.TimeLine.MainSequence;
            IEnumerable<Effect> mainEffects = sequence.Cast<Effect>();

            // handle deleted notes
            DeleteNotesFromSlideAnimationPane(notesDeleted, slide);

            // handle newly inserted notes
            AppendNotesToSlideAnimationPane(notesInserted, slide);

            // reorder notes on notes page
            ReorderNotesOnSlideAnimationPane(notes, slide, Microsoft.Office.Core.MsoTriState.msoFalse);
        }

        public static void ReorderAnimationsForCalloutsOnSlide(IntermediateResultTable intermediateResult, PowerPointSlide slide)
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
                if (nextClick == 0 && effectForClick == null)
                {
                    effectForClick = sequence.FindFirstAnimationForClick(1);
                    nextClick++;
                    prevClick++;
                }
                if (effectForClick != null && animeEffect.Shape.Name == effectForClick.Shape.Name)
                {
                    nextClick++;
                    prevClick++;
                }

                Shape shape = animeEffect.Shape;

                if (!shapeSet.Contains(shape))
                {
                    if (!shape.Name.Contains("PPTLabs Callout") || animeEffect.Exit != Microsoft.Office.Core.MsoTriState.msoTrue)
                    {
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

        private static void DeleteNotesFromSlideAnimationPane(List<Tuple<NameTag, string>> notesDeleted, PowerPointSlide slide)
        {
            for (int i = 0; i < notesDeleted.Count; i++)
            {
                List<Shape> shapes = slide.GetShapeWithName("PPTLabs Callout " + notesDeleted[i].Item1.Contents);
                slide.RemoveAnimationsForShapes(shapes);
            }
        }

        private static void AppendNotesToSlideAnimationPane(List<Tuple<NameTag, string>> notesInserted, PowerPointSlide slide)
        {
            int clickNo = 0;
            Sequence sequence = slide.TimeLine.MainSequence;
            Effect effect = sequence.FindFirstAnimationForClick(clickNo++);
            Shape prevShape = null;
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

        private static void ReorderNotesOnSlideAnimationPane(IEnumerable<NameTag> notes, PowerPointSlide slide,
            Microsoft.Office.Core.MsoTriState isExit = Microsoft.Office.Core.MsoTriState.msoFalse)
        {
            IEnumerable<Effect> mainEffects = slide.TimeLine.MainSequence.Cast<Effect>();
            Dictionary<int, Effect> appearIdxToEffect = new Dictionary<int, Effect>();
            string namescope = "PPTLabs Callout ";
            HashSet<string> shapes = new HashSet<string>();
            int count = mainEffects.Count();
            int notesCount = notes.Count();
            Dictionary<string, int> appearNameToIdx = new Dictionary<string, int>();
            Dictionary<int, Effect> disappearIdxToEffect = new Dictionary<int, Effect>();
            List<Effect> disappearEffectsToRemove = new List<Effect>();

            for (int i = 0; i < mainEffects.Count(); i++)
            {

                Effect animeEffect = mainEffects.ElementAt(i);
                Logger.Log("effect name is " + animeEffect.Shape.Name);
                int idx = shapes.Count();
                Shape shape = animeEffect.Shape;
                if (IsTargetedShapeEffect(animeEffect, namescope, isExit) && idx < notesCount)
                {
                    string tag = notes.ElementAt(idx).Contents;                                     
                    shapes.Add(tag);
                   
           
                    if (animeEffect.Shape.Name != namescope + tag)
                    {
                        Tuple<Effect, int> tuple = GetFirstEffectWithShapeNameAndCriteria(tag, namescope, i, slide, isExit);
                        if (tuple != null)
                        {
                            appearIdxToEffect[i] = tuple.Item1;
                            appearNameToIdx[tuple.Item1.Shape.Name] = i;
                            Logger.Log("moving "+tuple.Item1.Shape.Name + "to "+ (i + 1));
                            tuple.Item1.MoveTo(i + 1);
                            Logger.Log("moving " + animeEffect.Shape.Name + "to " + (i + 1));
                            animeEffect.MoveTo(tuple.Item2 + 1);
                        }
                    }
                    else
                    {
                        appearNameToIdx[animeEffect.Shape.Name] = i;
                        appearIdxToEffect[i] = animeEffect;
                    }
                }
                else
                {
                    appearIdxToEffect[i] = animeEffect;
                }
            }

            for (int i = 0; i < mainEffects.Count(); i++)
            {
                Effect effect = mainEffects.ElementAt(i);
                if (IsTargetedShapeEffect(effect, namescope, Microsoft.Office.Core.MsoTriState.msoTrue))
                {
                    string disappearEffectName = effect.Shape.Name;
                    if (appearNameToIdx.ContainsKey(disappearEffectName) && !disappearIdxToEffect.ContainsKey(appearNameToIdx[disappearEffectName]))
                    {
                        disappearIdxToEffect[appearNameToIdx[disappearEffectName]] = effect;
                    }
                    else
                    {
                        disappearEffectsToRemove.Add(effect);
                    }
                }
            }

            string prevKey = null;
            foreach (NameTag note in notes)
            {
                string currKey = namescope + note.Contents;
                if (prevKey != null)
                {
                    Effect _effect = appearIdxToEffect[appearNameToIdx[currKey]];
                    List<Shape> _shapes = slide.GetShapeWithName(prevKey);
                    if (_shapes.Count != 0)
                    {
                        int _key = appearNameToIdx[prevKey];
                        Effect effect = disappearIdxToEffect.ContainsKey(_key) ? disappearIdxToEffect[_key] :
                            slide.TimeLine.MainSequence.AddEffect(_shapes[0], MsoAnimEffect.msoAnimEffectAppear,
                            trigger: MsoAnimTriggerType.msoAnimTriggerAfterPrevious);
                        effect.Exit = Microsoft.Office.Core.MsoTriState.msoTrue;
                        effect.MoveAfter(_effect);
                       // Logger.Log("move effect with name " + effect.Shape.Name + " after appearance " + _effect.Shape.Name);
                    }
                }
                prevKey = currKey;
            }
            if (disappearIdxToEffect.ContainsKey(appearNameToIdx[prevKey]))
            {
                disappearEffectsToRemove.Add(disappearIdxToEffect[appearNameToIdx[prevKey]]);
            }
            
            foreach (Effect effect in disappearEffectsToRemove)
            {
              //  Logger.Log("removing disappear effect with name " + effect.Shape.Name);
                effect.Delete();
            }
        }

        private static Tuple<Effect, int> GetFirstEffectWithShapeNameAndCriteria(string tag, string namescope, int idx,
            PowerPointSlide slide, Microsoft.Office.Core.MsoTriState isExit)
        {
            IEnumerable<Effect> effects = slide.TimeLine.MainSequence.Cast<Effect>();
            int count = effects.Count();
            for (int i = idx + 1; i < count; i++)
            {
                
                Effect effect = effects.ElementAt(i);
                if (IsTargetedShapeEffect(effect, namescope, isExit) && effect.Shape.Name == namescope + tag)
                {           
                    return new Tuple<Effect, int>(effect, i);
                }
            }
            List<Shape> shape = slide.GetShapeWithName(namescope + tag);
            if (shape.Count != 0)
            {
                Effect effect = slide.TimeLine.MainSequence.AddEffect(shape[0], MsoAnimEffect.msoAnimEffectAppear);
                return new Tuple<Effect, int>(effect, count);
            }
            return null;
        }

        private static bool IsTargetedShapeEffect(Effect effect, string name, Microsoft.Office.Core.MsoTriState isExit)
        {
            return effect.Shape.Name.Contains(name) && effect.Exit == isExit;
        }

        private static Shape FindPPTLabsShapeForEffect(Effect effect)
        {
            if (effect == null)
            {
                return null;
            }
            Shape shape = effect.Shape;
            if (effect.Shape.Name.Contains("PPTLabs Callout "))
            {
                return shape;
            }
            return null;
        }
    }
}
