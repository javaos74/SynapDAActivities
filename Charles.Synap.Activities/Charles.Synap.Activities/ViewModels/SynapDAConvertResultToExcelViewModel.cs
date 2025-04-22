using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Activities.DesignViewModels;
using System.Net;
using System.Security.Cryptography;
using System.Activities;
using UiPath.Platform.ResourceHandling;

namespace Charles.Synap.Activities.ViewModels
{
    public class SynapDAConvertResultToExcelViewModel : DesignPropertiesViewModel
    {
        public DesignInArgument<IResource> ResultZip { get; set; }
        public DesignInArgument<string> ResultExcelFile { get; set; }
        public DesignOutArgument<int> TableCount { get; set; }
        public DesignOutArgument<string> ErrorMessage { get; set; }

        public SynapDAConvertResultToExcelViewModel(IDesignServices services) : base(services)
        {
        }

        protected override void InitializeModel()
        {
            /*
             * The base call will initialize the properties of the view model with the values from the xaml or with the default values from the activity
             */
            base.InitializeModel();

            PersistValuesChangedDuringInit(); // mandatory call only when you change the values of properties during initialization
            int propertyOrderIndex = 1;

            ResultZip.OrderIndex = propertyOrderIndex++;
            ResultExcelFile.OrderIndex = propertyOrderIndex++;
            TableCount.OrderIndex = propertyOrderIndex++;
            ErrorMessage.OrderIndex = propertyOrderIndex++;
        }
    }
}
