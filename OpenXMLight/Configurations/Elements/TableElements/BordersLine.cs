using OpenXMLight.Configurations.Formatting;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Text;
using System.Threading.Tasks;

using OpenXml = DocumentFormat.OpenXml.Wordprocessing;

namespace OpenXMLight.Configurations.Elements.TableElements
{
    public class BordersLine : Borders, IDisposable
    {
        #region Privaty properties
        private Borders top = new();
        private Borders bottom = new();
        private Borders left = new();
        private Borders right = new();
        private Borders insideH = new();
        private Borders insideV = new();

        private bool _isSubscribed = false;
        private bool _isDisposed = false;
        #endregion


        public override BordersType LineType 
        { 
            get => base.LineType;
            set => base.LineType = value; 
        }
        public override double LineWidth 
        { 
            get => base.LineWidth;
            set => base.LineWidth = value; 
        }


        internal OpenXml.TableBorders elementXml;

        public Borders Top { get => top; }
        public Borders Bottom { get => bottom; }
        public Borders Left { get => left; }
        public Borders Right {  get => right; }
        public Borders InsideHorizontal { get => insideH; }
        public Borders InsideVertical { get => insideV; }



        public BordersLine() : this(new OpenXml.TableBorders())
        {

        }
        internal BordersLine(OpenXml.TableBorders tblBorders)
        {
            elementXml = tblBorders;

            EnsureSubscribeEvent();

            //Top
            if (tblBorders.TopBorder != null)
            {
                top.LineType = BordersType.Parse(tblBorders.TopBorder?.Val);
                top.LineWidth = double.Parse(tblBorders.TopBorder?.Size) / Configuration.LineWidthInTable;
            }
            else
            {
                tblBorders.TopBorder = new OpenXml.TopBorder();
                top.LineType = BordersType.Single;
                top.LineWidth = 0.5;
            }
            //Bottom
            if (tblBorders.BottomBorder != null)
            {
                bottom.LineType = BordersType.Parse(tblBorders.BottomBorder?.Val);
                bottom.LineWidth = double.Parse(tblBorders.BottomBorder?.Size) / Configuration.LineWidthInTable;
            }
            else
            {
                tblBorders.BottomBorder = new OpenXml.BottomBorder();
                bottom.LineType = BordersType.Single;
                bottom.LineWidth = 0.5;
            }
            
            //Left
            if (tblBorders.LeftBorder != null)
            {
                left.LineType = BordersType.Parse(tblBorders.LeftBorder?.Val);
                left.LineWidth = double.Parse(tblBorders.LeftBorder?.Size) / Configuration.LineWidthInTable;
            }
            else
            {
                tblBorders.LeftBorder = new OpenXml.LeftBorder();
                left.LineType = BordersType.Single;
                left.LineWidth = 0.5;
            }

            //Right
            if (tblBorders.RightBorder != null)
            {
                right.LineType = BordersType.Parse(tblBorders.RightBorder?.Val);
                right.LineWidth = double.Parse(tblBorders.RightBorder?.Size) / Configuration.LineWidthInTable;
            }
            else
            {
                tblBorders.RightBorder = new OpenXml.RightBorder();
                right.LineType = BordersType.Single;
                right.LineWidth = 0.5;
            }
            
            //InsideH
            if (tblBorders.InsideHorizontalBorder != null)
            {
                insideH.LineType = BordersType.Parse(tblBorders.InsideHorizontalBorder?.Val);
                insideH.LineWidth = double.Parse(tblBorders.InsideHorizontalBorder.Size) / Configuration.LineWidthInTable;
            }
            else
            {
                tblBorders.InsideHorizontalBorder = new OpenXml.InsideHorizontalBorder();
                insideH.LineType = BordersType.Single;
                insideH.LineWidth = 0.5;
            }
            
            //InsideV
            if (tblBorders.InsideVerticalBorder != null)
            {
                insideV.LineType = BordersType.Parse(tblBorders.InsideVerticalBorder?.Val);
                insideV.LineWidth = double.Parse(tblBorders.InsideVerticalBorder?.Size) / Configuration.LineWidthInTable;
            }
            else
            {
                tblBorders.InsideVerticalBorder = new OpenXml.InsideVerticalBorder();
                insideV.LineType = BordersType.Single;
                insideV.LineWidth = 0.5;
            }
        }



        internal void OnBorderPropertyChanged(object sender, PropertyChangedEventArgs e)
        {
            if (_isDisposed) return;

            if (elementXml == null)
                throw new ArgumentNullException("Не обнаружены границы");

            var border = (Borders)sender;

            if (ReferenceEquals(border, top))
            {
                elementXml.TopBorder.Size = Convert.ToUInt32(top.LineWidth * Configuration.LineWidthInTable);
                elementXml.TopBorder.Val = top.LineType.Value;
            }
            else if (ReferenceEquals(border, bottom))
            {
                elementXml.BottomBorder.Size = Convert.ToUInt32(bottom.LineWidth * Configuration.LineWidthInTable);
                elementXml.BottomBorder.Val = bottom.LineType.Value;
            }
            else if (ReferenceEquals(border, left))
            {
                elementXml.LeftBorder.Size = Convert.ToUInt32(left.LineWidth * Configuration.LineWidthInTable);
                elementXml.LeftBorder.Val = left.LineType.Value;
            }
            else if (ReferenceEquals(border, right))
            {
                elementXml.RightBorder.Size = Convert.ToUInt32(right.LineWidth * Configuration.LineWidthInTable);
                elementXml.RightBorder.Val = right.LineType.Value;
            }
            else if (ReferenceEquals(border, insideH))
            {
                elementXml.InsideHorizontalBorder.Size = Convert.ToUInt32(insideH.LineWidth * Configuration.LineWidthInTable);
                elementXml.InsideHorizontalBorder.Val = insideH.LineType.Value;
            }
            else if (ReferenceEquals(border, insideV))
            {
                elementXml.InsideVerticalBorder.Size = Convert.ToUInt32(insideV.LineWidth * Configuration.LineWidthInTable);
                elementXml.InsideVerticalBorder.Val = insideV.LineType.Value;
            }
            else if (ReferenceEquals(border, this))
            {
                //Top
                elementXml.TopBorder.Size = Convert.ToUInt32(this.LineWidth * Configuration.LineWidthInTable);
                elementXml.TopBorder.Val = this.LineType.Value;

                //Bottom
                elementXml.BottomBorder.Size = Convert.ToUInt32(this.LineWidth * Configuration.LineWidthInTable);
                elementXml.BottomBorder.Val = this.LineType.Value;

                //Left
                elementXml.LeftBorder.Size = Convert.ToUInt32(this.LineWidth * Configuration.LineWidthInTable);
                elementXml.LeftBorder.Val = this.LineType.Value;

                //Right
                elementXml.RightBorder.Size = Convert.ToUInt32(this.LineWidth * Configuration.LineWidthInTable);
                elementXml.RightBorder.Val = this.LineType.Value;

                //InsideH
                elementXml.InsideHorizontalBorder.Size = Convert.ToUInt32(this.LineWidth * Configuration.LineWidthInTable);
                elementXml.InsideHorizontalBorder.Val = this.LineType.Value;

                //InsideV
                elementXml.InsideVerticalBorder.Size = Convert.ToUInt32(this.LineWidth * Configuration.LineWidthInTable);
                elementXml.InsideVerticalBorder.Val = this.LineType.Value;
            }
        }
        

        private void EnsureSubscribeEvent()
        {
            if (!_isSubscribed && !_isDisposed)
            {
                SubscribeEvent();
                _isSubscribed = true;
            }
        }
        private void SubscribeEvent()
        {
            if (_isDisposed) return;

            top.PropertyChanged += OnBorderPropertyChanged;
            bottom.PropertyChanged += OnBorderPropertyChanged;
            left.PropertyChanged += OnBorderPropertyChanged;
            right.PropertyChanged += OnBorderPropertyChanged;
            insideH.PropertyChanged += OnBorderPropertyChanged;
            insideV.PropertyChanged += OnBorderPropertyChanged;
            this.PropertyChanged += OnBorderPropertyChanged;
        }
        private void UnsubscribeEvent()
        {
            if (!_isSubscribed || _isDisposed) return;

            top.PropertyChanged -= OnBorderPropertyChanged;
            bottom.PropertyChanged -= OnBorderPropertyChanged;
            left.PropertyChanged -= OnBorderPropertyChanged;
            right.PropertyChanged -= OnBorderPropertyChanged;
            insideH.PropertyChanged -= OnBorderPropertyChanged;
            insideV.PropertyChanged -= OnBorderPropertyChanged;
            this.PropertyChanged -= OnBorderPropertyChanged;

            _isSubscribed = false;
        }

        #region IDisposable
        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }

        protected virtual void Dispose(bool disposing)
        {
            if (!_isDisposed)
            {
                if(disposing)
                {
                    UnsubscribeEvent();
                }

                _isDisposed = true;
            }
        }

        ~BordersLine() => Dispose(false);
        #endregion

        public static implicit operator OpenXml.TableBorders?(BordersLine borders) => borders?.elementXml;
    }
}
