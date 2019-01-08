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
            DeleteNotesFromSlideAnimationPane(notes, slide);

            // handle newly inserted notes
            AppendNotesToSlideAnimationPane(notesInserted, slide);

            // reorder notes on notes page
            ReorderNotesOnSlideAnimationPane(notes, slide, Microsoft.Office.Core.MsoTriState.msoFalse);
        }

        public static void AppendAnimationsForCalloutsToSlide(Shape shape, PowerPointSlide slide, bool byClick)
        {
            if (byClick)
            {
                slide.TimeLine.MainSequence.AddEffect(shape, MsoAnimEffect.msoAnimEffectAppear,
                    trigger: MsoAnimTriggerType.msoAnimTriggerOnPageClick);
            }
            else
            {
                slide.TimeLine.MainSequence.AddEffect(shape, MsoAnimEffect.msoAnimEffectAppear, trigger: MsoAnimTriggerType.msoAnimTriggerAfterPrevious);
            }
        }

        private static void DeleteNotesFromSlideAnimationPane(IEnumerable<NameTag> notes, PowerPointSlide slide)
        {
            List<Shape> shapes = slide.GetShapesWithPrefix("PPTLabs Callout ");
            foreach (Shape shape in shapes)
            {
                if (!notes.Any((nametag) => shape.Name.Contains(nametag.Contents)))
                {
                    shape.Delete();
                }
            }
        }

        private static void AppendNotesToSlideAnimationPane(List<Tuple<NameTag, string>> notesInserted, PowerPointSlide slide)
        {
            Sequence sequence = slide.TimeLine.MainSequence;
            Shape prevShape = null;
            bool isAnimationListEmpty = sequence.Cast<Effect>().Count() == 0;
            for (int i = 0; i < notesInserted.Count; i++)
            {
                List<Shape> shapes = slide.GetShapeWithName("PPTLabs Callout " + notesInserted[i].Item1.Contents);
                if (shapes.Count == 0)
                {
                    continue;
                }
                Shape currShape = shapes[0];
                if (!IsEffectExists(sequence.Cast<Effect>(), currShape.Name))
                {
                    Effect showEffect = slide.TimeLine.MainSequence.AddEffect(currShape, MsoAnimEffect.msoAnimEffectAppear); 
                    
                    if (prevShape != null)
                    {
                        Effect _effect = slide.TimeLine.MainSequence.AddEffect(prevShape, MsoAnimEffect.msoAnimEffectAppear,
                           trigger: MsoAnimTriggerType.msoAnimTriggerAfterPrevious);
                        _effect.Exit = Microsoft.Office.Core.MsoTriState.msoTrue;
                        _effect.MoveAfter(showEffect);
                    }
                }
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
            if (notesCount == 0)
            {
                return;
            }
            Dictionary<string, int> appearNameToIdx = new Dictionary<string, int>();
            Dictionary<int, Effect> disappearIdxToEffect = new Dictionary<int, Effect>();
            List<Effect> disappearEffectsToRemove = new List<Effect>();

            for (int i = 0; i < mainEffects.Count(); i++)
            {

                Effect animeEffect = mainEffects.ElementAt(i);
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
                            tuple.Item1.MoveTo(i + 1);
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

        private static bool IsEffectExists(IEnumerable<Effect> effects, string name)
        {
            return effects.Any((effect) => effect.Shape.Name == name);
        }
    }
}
