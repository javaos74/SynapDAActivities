using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Activities.DesignViewModels;
using System.Activities;
using System.Data;

namespace Charles.Synap.Activities.ViewModels
{
    public class SynapDAConvertExcelToJsonViewModel : DesignPropertiesViewModel
    {
        public DesignInArgument<DataTable> InputDataTable { get; set; }
        public DesignInArgument<string> OutputJsonFilePath { get; set; }
        public DesignOutArgument<int> RecordCount { get; set; }
        public DesignOutArgument<string> ErrorMessage { get; set; }

        public SynapDAConvertExcelToJsonViewModel(IDesignServices services) : base(services)
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

            InputDataTable.OrderIndex = propertyOrderIndex++;
            OutputJsonFilePath.OrderIndex = propertyOrderIndex++;
            RecordCount.OrderIndex = propertyOrderIndex++;
            ErrorMessage.OrderIndex = propertyOrderIndex++;
        }
    }
}