using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace FileProcessor.Models
{
    public class TextContent
    {
        private string fontName;
        private string fontSize;
        private string text;

        public string FontName
        {
            get
            {
                return fontName;
            }

            set
            {
                fontName = value;
            }
        }

        public string FontSize
        {
            get
            {
                return fontSize;
            }

            set
            {
                fontSize = value;
            }
        }

        public string Text
        {
            get
            {
                return text;
            }

            set
            {
                text = value;
            }
        }
        public TextContent(string fontName, string fontSize, string text)
        {
            this.fontName = fontName;
            this.fontSize = fontSize;
            this.text = text;
        }
    }
}