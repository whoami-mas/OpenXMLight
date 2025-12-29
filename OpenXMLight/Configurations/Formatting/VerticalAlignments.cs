using DocumentFormat.OpenXml;
using OpenXMLight.config;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using OpenXml = DocumentFormat.OpenXml.Wordprocessing;

namespace OpenXMLight.Configurations.Formatting
{
    public readonly record struct VerticalAlignments : IEnumValue<OpenXml.TableVerticalAlignmentValues>
    {
        public OpenXml.TableVerticalAlignmentValues Value => _value ?? OpenXml.TableVerticalAlignmentValues.Top;



        public static VerticalAlignments Top => new VerticalAlignments(OpenXml.TableVerticalAlignmentValues.Top);
        public static VerticalAlignments Center => new VerticalAlignments(OpenXml.TableVerticalAlignmentValues.Center);
        public static VerticalAlignments Bottom => new VerticalAlignments(OpenXml.TableVerticalAlignmentValues.Bottom);



        private readonly OpenXml.TableVerticalAlignmentValues? _value;
        public VerticalAlignments(OpenXml.TableVerticalAlignmentValues value)
        {
            _value = value;
        }

        internal static VerticalAlignments Parse(OpenXml.TableVerticalAlignmentValues value)
        {
            return value switch
            {
                var v when v == OpenXml.TableVerticalAlignmentValues.Top => VerticalAlignments.Top,
                var v when v == OpenXml.TableVerticalAlignmentValues.Center => VerticalAlignments.Center,
                var v when v == OpenXml.TableVerticalAlignmentValues.Bottom => VerticalAlignments.Bottom,
                
                _ => VerticalAlignments.Top
            };
        }
    }
}
