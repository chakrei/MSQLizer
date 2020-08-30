using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Web;

namespace FileToDB.Models
{
    public static class AppKeyHelper
    {
        private static string acceptedFileTypes;
        public static string AcceptedFileTypes
        {
            get
            {
                if (string.IsNullOrEmpty(acceptedFileTypes))
                    acceptedFileTypes = ConfigurationManager.AppSettings["acceptedFileTyps"];
                return acceptedFileTypes;
            }
        }
    }
}