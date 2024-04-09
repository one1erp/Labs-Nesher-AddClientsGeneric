using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Runtime.InteropServices;
using AddClientFromExcel;
using Common;
using CopyAddresses;
using LSEXT;
using LSSERVICEPROVIDERLib;

namespace AddClientsGeneric
{


    [ComVisible(true)]
    [ProgId("AddClientsGeneric.AddClientsGenericcls")]
    public class AddClientsGenericcls : IGenericExtension
    {
        private INautilusServiceProvider sp;

        public ExecuteExtension CanExecute(ref IExtensionParameters Parameters)
        {
            return ExecuteExtension.exEnabled;
        }

        public void Execute(ref LSExtensionParameters Parameters)
        {
            sp = Parameters["SERVICE_PROVIDER"];
            var ntlsCon = Utils.GetNtlsCon(sp);
            Utils.CreateConstring(ntlsCon);
                var aForm1 = new Form1();
             aForm1.Show();
            //CopyAddressesForm a = new CopyAddressesForm();
            //a.Show();






        }
    }
}
