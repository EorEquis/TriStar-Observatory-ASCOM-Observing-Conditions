using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Utils
{
    /// <summary>
    /// Subclass of WebClient to provide access to the timeout property
    /// </summary>
    public partial class WebClient : System.Net.WebClient
    {

        private int _TimeoutMS = 1000;

        public WebClient() : base()
        {
        }
        public WebClient(int TimeoutMS) : base()
        {
            _TimeoutMS = TimeoutMS;
        }
        /// <summary>
        /// Set the web call timeout in Milliseconds
        /// </summary>
        /// <value></value>
        public int setTimeout
        {
            set
            {
                _TimeoutMS = value;
            }
        }


        protected override System.Net.WebRequest GetWebRequest(Uri address)
        {
            System.Net.WebRequest w = base.GetWebRequest(address);
            if (_TimeoutMS != 0)
            {
                w.Timeout = _TimeoutMS;
            }
            return w;
        }

    }
}
