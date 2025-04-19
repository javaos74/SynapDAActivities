using System.Activities;
using System.Activities.DesignViewModels;

namespace Charles.Synap.Activities.ViewModels
{
    public class SynapDARequestViewModel : DesignPropertiesViewModel
    {
        /*
         * The result property comes from the activity's base class
         */
        public DesignInArgument<string> Endpoint { get; set; }
        public DesignInArgument<string> ApiKey { get; set; }
        public DesignInArgument<string> InputFilePath { get; set; }
        public DesignOutArgument<string> FID { get; set; }

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
            InputFilePath.OrderIndex = propertyOrderIndex++;
            FID.OrderIndex = propertyOrderIndex++;

        }
    }
}
