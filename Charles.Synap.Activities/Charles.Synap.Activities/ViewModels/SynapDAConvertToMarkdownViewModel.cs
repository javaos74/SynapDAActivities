using System;
using System.Activities;
using System.Activities.DesignViewModels;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Charles.Synap.Activities.ViewModels
{
    public class SynapDAConvertToMarkdownViewModel : DesignPropertiesViewModel
    {
        public DesignInArgument<string> ResultZipFIle { get; set; }
        public DesignOutArgument<int> PageCount { get; set; }
        public DesignOutArgument<string> ErrorMessage { get; set; }
        public DesignOutArgument<string> MarkdownBody { get; set; }

        public SynapDAConvertToMarkdownViewModel(IDesignServices services) : base(services)
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

            ResultZipFIle.OrderIndex = propertyOrderIndex++;
            MarkdownBody.OrderIndex = propertyOrderIndex++;
            PageCount.OrderIndex = propertyOrderIndex++;
            ErrorMessage.OrderIndex = propertyOrderIndex++;
        }

    }
}
