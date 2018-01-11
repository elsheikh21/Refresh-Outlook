using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Linq;
using System.ServiceProcess;
using System.Text;
using System.Threading.Tasks;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace RefreshOutlook {
    public partial class refresh :ServiceBase {

        public refresh() {
            InitializeComponent();
        }

        protected override void OnStart(string[] args) {
            refreshInbox();
        }

        public static void refreshInbox() {
            // create new Outlook App
            Outlook.Application myApp = new Outlook.Application(); 
            Outlook.NameSpace mapiNameSpace = myApp.GetNamespace( "MAPI" );
            Outlook.MAPIFolder myInbox = mapiNameSpace.GetDefaultFolder( Outlook.OlDefaultFolders.olFolderInbox );
            // Syncying the folder
            Outlook._SyncObject _syncObject = null;
            _syncObject = mapiNameSpace.SyncObjects[1];
            _syncObject.Start();
            System.Threading.Thread.Sleep( 5000 );
            _syncObject.Stop();
            _syncObject = null;
        }

        protected override void OnStop() {

        }

    }
}
