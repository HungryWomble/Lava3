using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.Serialization;
using System.ServiceModel;
using System.ServiceModel.Web;
using System.Text;

namespace Lava3.Rest
{
    // NOTE: You can use the "Rename" command on the "Refactor" menu to change the interface name "IService1" in both code and config file together.
    [ServiceContract]
    public interface IAccounts
    {

        [OperationContract]
        void ProcessCurrentAccount(string filename);
        [OperationContract]
        void ProcessCreditCard(string filename);

    }

    
}
