using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Web.Services;
using System.Net.Http;
using System.Net;
using Microsoft.Identity.Client;
using System.Threading.Tasks;
using System.Web.Http;

namespace CodeCrasher.Web.Microsoft
{
    public partial class CodeCrasherWebForm : System.Web.UI.Page
    {
        protected void Page_Load(object sender, EventArgs e)
        {
            Page.ClientScript.RegisterClientScriptInclude("CodeCrasherWebForm.js", Page.ResolveUrl("~/pages/js/CodeCrasherWebForm.js"));
        }

        [WebMethod()]
        public string GetProfile()
        {
            return "The profiles";
        }

    }
}