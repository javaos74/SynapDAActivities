using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Activities.DesignViewModels;
using System.Activities;

namespace Charles.Synap.Activities.ViewModels
{
    public class SynapDACheckFileStatusViewModel : DesignPropertiesViewModel
    {
        public DesignInArgument<string> Endpoint { get; set; }
        public DesignInArgument<string> ApiKey { get; set; }
        public DesignInArgument<string> FID { get; set; }
        public DesignOutArgument<string> FileStatus { get; set; }
        public DesignOutArgument<int> TotalPages { get; set; }
        public DesignOutArgument<int> ReturnCode { get; set; }
        public DesignOutArgument<int> Status { get; set; }
        public DesignOutArgument<string> ErrorMessage { get; set; }

        public SynapDACheckFileStatusViewModel(IDesignServices services) : base(services)
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

            Endpoint.OrderIndex = propertyOrderIndex++;
            ApiKey.OrderIndex = propertyOrderIndex++;
            FID.OrderIndex = propertyOrderIndex++;
            FileStatus.OrderIndex = propertyOrderIndex++;
            TotalPages.OrderIndex = propertyOrderIndex++;
            ReturnCode.OrderIndex = propertyOrderIndex++;
            Status.OrderIndex = propertyOrderIndex++;
            ErrorMessage.OrderIndex = propertyOrderIndex++;
        }
    }
}