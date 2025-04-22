using System.Activities;
using System.Activities.DesignViewModels;
using UiPath.Platform.ResourceHandling;

namespace Charles.Synap.Activities.ViewModels
{
    public class SynapDARequestViewModel : DesignPropertiesViewModel
    {
        /*
         * The result property comes from the activity's base class
         */
        public DesignInArgument<string> Endpoint { get; set; }
        public DesignInArgument<string> ApiKey { get; set; }
        public DesignInArgument<IResource> InputFile { get; set; }
        public DesignOutArgument<string> FID { get; set; }
        public DesignOutArgument<string> ErrorMessage { get; set; }
        public DesignOutArgument<int> Status { get; set; }

        public SynapDARequestViewModel(IDesignServices services) : base(services)
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
            InputFile.OrderIndex = propertyOrderIndex++;
            FID.OrderIndex = propertyOrderIndex++;
            Status.OrderIndex = propertyOrderIndex++;
            ErrorMessage.OrderIndex = propertyOrderIndex++;

        }
    }
}
