using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.Serialization;
using System.ServiceModel;
using System.ServiceModel.Web;
using System.Text;

namespace Lava3.Rest
{
    public class Service1 : IAccounts
    {
        public void ProcessCreditCard(string filename)
        {
            throw new NotImplementedException();
        }

        public void ProcessCurrentAccount(string filename)
        {
            throw new NotImplementedException();
        }
        
    }
}
