using System;
using System.Collections.Generic;
using System.Text;

namespace GraphicsNET
{
    class Data1C
    {
        public static object Object1C
        {
            get
            {
                return m_Object1C;
            }
            set
            {
                m_Object1C = value;                
            }
        }
       
        private static object m_Object1C;        
    }
}
